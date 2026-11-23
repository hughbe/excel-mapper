using System.ComponentModel.DataAnnotations;
using System.Linq;

namespace ExcelMapper.Tests;

public class ExcelMappingErrorTests
{
    public class Row
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }

    // Copies what each error says, as an error is only valid during Add.
    private sealed class RecordingSink : IExcelErrorSink
    {
        public List<(ExcelMappingErrorKind Kind, int RowIndex, int ColumnIndex, string? ColumnName, string? Member, string? Text, bool IsEmpty, ExcelSheet? Sheet, ExcelMappingException Exception)> Errors { get; } = [];

        public void Add(in ExcelMappingError error)
            => Errors.Add((error.Kind, error.RowIndex, error.ColumnIndex, error.ColumnName, error.Member?.Name, error.Value.Text, error.Value.IsEmpty, error.Sheet, error.ToException()));
    }

    [Fact]
    public void Properties_InvalidValue_DescribeCell()
    {
        using var importer = Helpers.GetCsvImporter("Name,Age\nAlice,abc\n");
        var sheet = importer.ReadSheet();
        var sink = new RecordingSink();

        Assert.Empty(sheet.ReadRows<Row>(sink));

        var error = Assert.Single(sink.Errors);
        Assert.Equal(ExcelMappingErrorKind.InvalidValue, error.Kind);
        Assert.Equal(1, error.RowIndex);
        Assert.Equal(1, error.ColumnIndex);
        Assert.Equal("Age", error.ColumnName);
        Assert.Equal("Age", error.Member);
        Assert.Equal("abc", error.Text);
        Assert.Same(sheet, error.Sheet);
    }

    [Fact]
    public void Properties_EmptyValue_DescribeCell()
    {
        using var importer = Helpers.GetCsvImporter("Name,Age\nAlice,\n");
        var sink = new RecordingSink();

        Assert.Empty(importer.ReadSheet().ReadRows<Row>(sink));

        var error = Assert.Single(sink.Errors);
        Assert.Equal(ExcelMappingErrorKind.EmptyValue, error.Kind);
        Assert.True(error.IsEmpty);
        Assert.Equal("B2", error.Exception.CellAddress);
    }

    [Fact]
    public void ToException_CellErrors_SameAsReadRowResults()
    {
        // The exception a sink creates is the one ReadRowResults collects.
        const string Csv = "Name,Age\nAlice,abc\nBob,\nCarol,99999999999\n";
        using var resultsImporter = Helpers.GetCsvImporter(Csv);
        var expected = resultsImporter.ReadSheet().ReadRowResults<Row>().SelectMany(r => r.Errors).ToArray();

        using var importer = Helpers.GetCsvImporter(Csv);
        var sink = new RecordingSink();
        Assert.Empty(importer.ReadSheet().ReadRows<Row>(sink));

        Assert.Equal(expected.Select(e => e.Message), sink.Errors.Select(e => e.Exception.Message));
        Assert.Equal(expected.Select(e => e.InnerException?.GetType()), sink.Errors.Select(e => e.Exception.InnerException?.GetType()));
        Assert.Equal(expected.Select(e => e.Member), sink.Errors.Select(e => e.Exception.Member));
        Assert.Equal(expected.Select(e => e.CellAddress), sink.Errors.Select(e => e.Exception.CellAddress));
    }

    public record Person(string Name, int Age)
    {
        public int Age { get; } = Age >= 0 ? Age : throw new ArgumentOutOfRangeException(nameof(Age), "Age can't be negative.");
    }

    [Fact]
    public void Kind_ConstructorThrows_IsException()
    {
        using var importer = Helpers.GetCsvImporter("Name,Age\nAlice,-1\nBob,2\n");
        var sink = new RecordingSink();

        var people = importer.ReadSheet().ReadRows<Person>(sink).ToArray();

        Assert.Equal("Bob", Assert.Single(people).Name);
        var error = Assert.Single(sink.Errors);
        Assert.Equal(ExcelMappingErrorKind.Exception, error.Kind);
        Assert.Equal(1, error.RowIndex);
        Assert.Equal(-1, error.ColumnIndex);
        Assert.Null(error.Text);
        Assert.IsType<ArgumentOutOfRangeException>(error.Exception.InnerException);
    }

    public class AnnotatedRow
    {
        [Range(0, 120)]
        public int Age { get; set; }
    }

    [Fact]
    public void Kind_DataAnnotationFails_IsDataAnnotation()
    {
        using var importer = Helpers.GetCsvImporter("Age\n200\n30\n");
        importer.Configuration.ValidateDataAnnotations = true;
        var sink = new RecordingSink();

        Assert.Equal(30, Assert.Single(importer.ReadSheet().ReadRows<AnnotatedRow>(sink)).Age);

        var error = Assert.Single(sink.Errors);
        Assert.Equal(ExcelMappingErrorKind.DataAnnotation, error.Kind);
        Assert.IsType<ValidationException>(error.Exception.InnerException);
    }

    public class ListRow
    {
        public int[] Values { get; set; } = [];
    }

    [Fact]
    public void Kind_InvalidElement_IsInvalidValue()
    {
        using var importer = Helpers.GetCsvImporter("Values\n\"1,x,3\"\n");
        var sink = new RecordingSink();

        Assert.Empty(importer.ReadSheet().ReadRows<ListRow>(sink));

        var error = Assert.Single(sink.Errors);
        Assert.Equal(ExcelMappingErrorKind.InvalidValue, error.Kind);
        Assert.Equal("x", error.Text);
        Assert.Equal("Values", error.Member);
    }
}

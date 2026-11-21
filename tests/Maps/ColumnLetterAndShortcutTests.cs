using System.IO;
using System.Linq;
using System.Text;

namespace ExcelMapper.Tests;

public class ColumnLetterAndShortcutTests
{
    public class Row
    {
        [ExcelColumnLetter("C")]
        public string Third { get; set; } = "";

        [ExcelColumnLetter("a")]
        public string First { get; set; } = "";
    }

    [Fact]
    public void ExcelColumnLetter_MapsColumn()
    {
        using var importer = Helpers.GetCsvImporter("x,y,z\n1,2,3\n");
        var row = importer.ReadSheet().ReadRows<Row>().Single();
        Assert.Equal("3", row.Third);
        Assert.Equal("1", row.First);
    }

    [Theory]
    [InlineData("A", 0)]
    [InlineData("c", 2)]
    [InlineData(" Z ", 25)]
    [InlineData("AA", 26)]
    [InlineData("XFD", 16383)]
    public void ExcelColumnLetter_Ctor_SetsIndexAndLetter(string letter, int expectedIndex)
    {
        var attribute = new ExcelColumnLetterAttribute(letter);
        Assert.Equal(expectedIndex, attribute.Index);
        Assert.Equal(letter.Trim().ToUpperInvariant(), attribute.Letter);
    }

    [Theory]
    [InlineData("")]
    [InlineData("1")]
    [InlineData("A1")]
    [InlineData("XFE")]
    [InlineData("É")]
    public void ExcelColumnLetter_InvalidLetter_ThrowsArgumentException(string letter)
    {
        Assert.Throws<ArgumentException>("letter", () => new ExcelColumnLetterAttribute(letter));
    }

    [Fact]
    public void ExcelColumnLetter_Null_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("letter", () => new ExcelColumnLetterAttribute(null!));
    }

    public class FluentRow
    {
        public string Value { get; set; } = "";
    }

    [Fact]
    public void WithColumnLetter_MapsColumn()
    {
        using var importer = Helpers.GetCsvImporter("x,y\n1,2\n");
        importer.Configuration.RegisterClassMap<FluentRow>(c => c.Map(p => p.Value).WithColumnLetter("b"));
        Assert.Equal("2", importer.ReadSheet().ReadRows<FluentRow>().Single().Value);
    }

    [Fact]
    public void WithColumnLetter_Invalid_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<FluentRow>();
        Assert.Throws<ArgumentException>("columnLetter", () => map.Map(p => p.Value).WithColumnLetter("1"));
        Assert.Throws<ArgumentNullException>("columnLetter", () => map.Map(p => p.Value).WithColumnLetter(null!));
    }

    public record Person(string Name, int Age);

    [Fact]
    public void ReadRows_Path_ReadsFirstSheet()
    {
        var rows = ExcelImporter.ReadRows<string[]>(Helpers.GetResourcePath("Primitives.xlsx"));
        Assert.NotEmpty(rows);
    }

    [Fact]
    public void ReadRows_PathCsv_Success()
    {
        var rows = ExcelImporter.ReadRows<string[]>(Helpers.GetResourcePath("Primitives.csv"), ExcelImporterFileType.Csv);
        Assert.NotEmpty(rows);
    }

    [Fact]
    public void ReadRows_StreamWithConfiguration_Success()
    {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("Full Name,Age\nAlice,30\nBob,25\n"));
        var rows = ExcelImporter.ReadRows<Person>(stream, ExcelImporterFileType.Csv, c =>
            c.RegisterClassMap<Person>(m =>
            {
                m.Map(p => p.Name).WithColumnName("Full Name");
                m.Map(p => p.Age);
            }));

        Assert.Equal([new Person("Alice", 30), new Person("Bob", 25)], rows);
    }

    [Fact]
    public void ReadRows_Stream_ClosesStream()
    {
        // Matches disposing an importer created from a stream: ExcelDataReader closes the stream.
        var stream = new MemoryStream(Encoding.UTF8.GetBytes("Name,Age\nAlice,30\n"));
        Assert.Single(ExcelImporter.ReadRows<Person>(stream, ExcelImporterFileType.Csv));
        Assert.False(stream.CanRead);
    }

    [Fact]
    public void ReadRows_NullArguments_Throw()
    {
        Assert.Throws<ArgumentNullException>("path", () => ExcelImporter.ReadRows<Person>((string)null!));
        Assert.Throws<ArgumentNullException>("stream", () => ExcelImporter.ReadRows<Person>((Stream)null!));
    }
}

using System.Linq;
using System.Reflection;
using ExcelMapper.Abstractions;
using ExcelMapper.Mappers;

namespace ExcelMapper.Tests;

/// <summary>
/// Members whose mappers are all built-in typed converters (strings, numbers, dates, bools, GUIDs)
/// are mapped without boxing. These check that the typed path behaves exactly like the pipeline.
/// </summary>
public class TypedConverterTests
{
    [Fact]
    public void ReadRows_InvalidValue_CustomFallbackReceivesMappersException()
    {
        using var importer = Helpers.GetCsvImporter("Value\nabc\n");
        var fallback = new RecordingFallback();
        importer.Configuration.RegisterClassMap<IntRow>(c => c.Map(r => r.Value).WithInvalidFallbackItem(fallback));
        Assert.Equal(-1, importer.ReadSheet().ReadRows<IntRow>().Single().Value);
        Assert.IsType<FormatException>(fallback.Exception);
    }

    private sealed class RecordingFallback : IFallbackItem
    {
        public Exception? Exception { get; private set; }

        public object? PerformFallback(ExcelSheet sheet, int rowIndex, ReadCellResult readResult, Exception? exception, MemberInfo? member)
        {
            Exception = exception;
            return -1;
        }
    }

    [Fact]
    public void ReadRows_InvalidValue_ErrorUnchanged()
    {
        using var importer = Helpers.GetCsvImporter("Value\n2.5\n");
        var sheet = importer.ReadSheet();
        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<IntRow>().ToArray());
        Assert.Equal("A2", exception.CellAddress);
        Assert.Contains("Cannot assign \"2.5\" to member \"Value\"", exception.Message);
        // As before the typed path: int.Parse reports a fractional part as an overflow.
        Assert.IsType<OverflowException>(exception.InnerException);
    }

    [Fact]
    public void ReadRows_InvalidValueWithFixedFallback_ReturnsFallback()
    {
        using var importer = Helpers.GetCsvImporter("Value\nabc\n7\n");
        importer.Configuration.RegisterClassMap<IntRow>(c => c.Map(r => r.Value).WithInvalidFallback(-1));
        Assert.Equal([-1, 7], importer.ReadSheet().ReadRows<IntRow>().Select(r => r.Value));
    }

    [Fact]
    public void ReadRows_MappersChangedAfterReading_UsesNewMappers()
    {
        using var importer = Helpers.GetCsvImporter("Value\n1\n2\n");
        var map = new ExcelClassMap<IntRow>();
        var valueMap = map.Map(r => r.Value);
        importer.Configuration.RegisterClassMap(map);
        var sheet = importer.ReadSheet();
        Assert.Equal(1, sheet.ReadRow<IntRow>().Value);

        // A custom mapper isn't a typed converter: the member goes back to the pipeline.
        valueMap.WithMappers(new DoublingMapper());
        Assert.Equal(4, sheet.ReadRow<IntRow>().Value);
    }

    private sealed class DoublingMapper : ICellMapper
    {
        public CellMapperResult Map(ReadCellResult readResult) => CellMapperResult.Success(int.Parse(readResult.GetString()!) * 2);
    }

#if LIBRARY_NET8
    [Fact]
    public void ReadRows_CustomMapperThenBuiltIn_UsesPipeline()
    {
        // A custom mapper that ignores some values, then the built-in one.
        using var importer = Helpers.GetCsvImporter("Value\nn/a\n3\n");
        importer.Configuration.RegisterClassMap<IntRow>(c => c.Map(r => r.Value)
            .WithMappers(new IgnoreNotApplicable(), new INumberBaseMapper<int>())
            .WithInvalidFallback(0));
        Assert.Equal([0, 3], importer.ReadSheet().ReadRows<IntRow>().Select(r => r.Value));
    }

    private sealed class IgnoreNotApplicable : ICellMapper
    {
        public CellMapperResult Map(ReadCellResult readResult)
            => readResult.GetString() == "n/a" ? CellMapperResult.Invalid(new FormatException()) : CellMapperResult.Ignore();
    }
#endif

    [Fact]
    public void ReadRows_EmptyValues_UseEmptyFallbacks()
    {
        using var importer = Helpers.GetCsvImporter("Value,Name,When\n,,\n");
        importer.Configuration.RegisterClassMap<Mixed>(c =>
        {
            c.Map(r => r.Value).WithEmptyFallback(5);
            c.Map(r => r.Name);
            c.Map(r => r.When).WithEmptyFallback(new DateTime(2024, 1, 1));
        });
        var row = importer.ReadSheet().ReadRows<Mixed>().Single();
        Assert.Equal(5, row.Value);
        Assert.Null(row.Name);
        Assert.Equal(new DateTime(2024, 1, 1), row.When);
    }

    [Theory]
    [InlineData("yes", true)]
    [InlineData("No", false)]
    [InlineData("1", true)]
    [InlineData("0", false)]
    [InlineData("TRUE", true)]
    public void ReadRows_Bool_ReturnsExpected(string text, bool expected)
    {
        using var importer = Helpers.GetCsvImporter($"Value\n{text}\n");
        Assert.Equal(expected, importer.ReadSheet().ReadRows<BoolRow>().Single().Value);
    }

    [Fact]
    public void ReadRows_TypedCellsXlsx_ReturnsExpected()
    {
        var date = new DateTime(2024, 3, 4);
        using var importer = XlsxBuilder.Create(["Value", "Name", "When"], [2.0, "a", date], [3.0, "b", 45000.0]);
        var rows = importer.ReadSheet().ReadRows<Mixed>().ToArray();
        Assert.Equal([2, 3], rows.Select(r => r.Value));
        Assert.Equal(date, rows[0].When);
        Assert.Equal(new DateTime(2023, 3, 15), rows[1].When);
    }

    public class IntRow
    {
        public int Value { get; set; }
    }

    public class BoolRow
    {
        public bool Value { get; set; }
    }

    public class Mixed
    {
        public int Value { get; set; }
        public string? Name { get; set; }
        public DateTime When { get; set; }
    }
}

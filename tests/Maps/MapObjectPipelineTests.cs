using System.Linq;

namespace ExcelMapper.Tests;

/// <summary>
/// A member whose mappers aren't all typed converters (here, because of a mapping dictionary) is
/// mapped through the pipeline as objects. These check that the built-in conversions give the same
/// values on that path.
/// </summary>
public class MapObjectPipelineTests
{
    [Fact]
    public void ReadRows_Bool_ReturnsExpected()
    {
        var values = Read<bool>(m => m.WithMapping(new Dictionary<string, bool> { ["n/a"] = true }),
            true, "1", "0", "yes", "No", "true", "False", "n/a");
        Assert.Equal([true, true, false, true, false, true, false, true], values);
    }

    [Fact]
    public void ReadRows_DateTime_ReturnsExpected()
    {
        var date = new DateTime(2024, 3, 4, 10, 30, 0);
        var values = Read<DateTime>(m => m.WithFormats("yyyyMMdd").WithMapping(new Dictionary<string, DateTime> { ["n/a"] = DateTime.MinValue }),
            date, 45355.5, "20240304", "n/a");
        Assert.Equal([date, DateTime.FromOADate(45355.5), new DateTime(2024, 3, 4), DateTime.MinValue], values);
    }

    [Fact]
    public void ReadRows_DateTimeOffset_ReturnsExpected()
    {
        var date = new DateTime(2024, 3, 4, 10, 30, 0);
        var values = Read<DateTimeOffset>(m => m.WithFormats("yyyyMMdd").WithMapping(new Dictionary<string, DateTimeOffset> { ["n/a"] = DateTimeOffset.MinValue }),
            date, 45355.5, "20240304", "n/a");
        Assert.Equal([date, DateTime.FromOADate(45355.5), new DateTime(2024, 3, 4), DateTimeOffset.MinValue.DateTime], values.Select(v => v.DateTime));
    }

    [Fact]
    public void ReadRows_TimeSpan_ReturnsExpected()
    {
        var values = Read<TimeSpan>(m => m.WithFormats(@"hh\-mm").WithMapping(new Dictionary<string, TimeSpan> { ["n/a"] = TimeSpan.MaxValue }),
            new XlsxBuilder.Duration(TimeSpan.FromHours(30)), new TimeOnly(10, 30), 0.5, "01-30", "n/a");
        Assert.Equal([TimeSpan.FromHours(30), new TimeSpan(10, 30, 0), TimeSpan.FromHours(12), new TimeSpan(1, 30, 0), TimeSpan.MaxValue], values);
    }

#if LIBRARY_NET8
    [Fact]
    public void ReadRows_DateOnly_ReturnsExpected()
    {
        var values = Read<DateOnly>(m => m.WithFormats("yyyyMMdd").WithMapping(new Dictionary<string, DateOnly> { ["n/a"] = DateOnly.MinValue }),
            new DateOnly(2024, 3, 4), 45355.0, "20240304", "n/a");
        Assert.Equal([new DateOnly(2024, 3, 4), DateOnly.FromDateTime(DateTime.FromOADate(45355)), new DateOnly(2024, 3, 4), DateOnly.MinValue], values);
    }

    [Fact]
    public void ReadRows_TimeOnly_ReturnsExpected()
    {
        var values = Read<TimeOnly>(m => m.WithFormats("HH-mm").WithMapping(new Dictionary<string, TimeOnly> { ["n/a"] = TimeOnly.MinValue }),
            new TimeOnly(10, 30), 0.5, "14-15", "n/a");
        Assert.Equal([new TimeOnly(10, 30), new TimeOnly(12, 0), new TimeOnly(14, 15), TimeOnly.MinValue], values);
    }

    [Fact]
    public void ReadRows_InvalidParsable_ThrowsWithParseException()
    {
        using var importer = Helpers.GetCsvImporter("Value\nbad\n");
        importer.Configuration.RegisterClassMap<Row<Code>>(c => c.Map(r => r.Value));
        var exception = Assert.Throws<ExcelMappingException>(() => importer.ReadSheet().ReadRows<Row<Code>>().ToArray());
        Assert.IsType<FormatException>(exception.InnerException);
    }

    public readonly struct Code : IParsable<Code>
    {
        public string Value { get; init; }

        public static Code Parse(string s, IFormatProvider? provider)
            => s.Length == 2 ? new Code { Value = s } : throw new FormatException($"\"{s}\" isn't a code.");

        public static bool TryParse(string? s, IFormatProvider? provider, out Code result)
        {
            result = default;
            return false;
        }
    }
#endif

    [Fact]
    public void ReadRows_Guid_ReturnsExpected()
    {
        var guid = Guid.NewGuid();
        var values = Read<Guid>(m => m.WithMapping(new Dictionary<string, Guid> { ["n/a"] = Guid.Empty }), guid.ToString(), "n/a");
        Assert.Equal([guid, Guid.Empty], values);
    }

    [Fact]
    public void ReadRows_ConverterReturnsInvalidWithException_ThrowsWithException()
    {
        var exception = new InvalidOperationException("Not valid.");
        using var importer = Helpers.GetCsvImporter("Value\nx\n");
        importer.Configuration.RegisterClassMap<Row<int>>(c => c.Map(r => r.Value)
            .WithConverter(new InvalidConverter(exception))
            .WithMapping(new Dictionary<string, int> { ["n/a"] = 0 }));
        var thrown = Assert.Throws<ExcelMappingException>(() => importer.ReadSheet().ReadRows<Row<int>>().ToArray());
        Assert.Same(exception, thrown.InnerException);
    }

    [Fact]
    public void ReadRows_NullableConverterReturnsInvalidWithException_ThrowsWithException()
    {
        var exception = new InvalidOperationException("Not valid.");
        using var importer = Helpers.GetCsvImporter("Value\nx\n");
        importer.Configuration.RegisterClassMap<Row<int?>>(c => c.Map(r => r.Value).WithConverter(new InvalidConverter(exception)));
        var thrown = Assert.Throws<ExcelMappingException>(() => importer.ReadSheet().ReadRows<Row<int?>>().ToArray());
        Assert.Same(exception, thrown.InnerException);
    }

    private sealed class InvalidConverter(Exception exception) : ICellConverter<int>
    {
        public CellConversionResult<int> Convert(CellValue value, CellContext context) => CellConversionResult<int>.Invalid(exception);
    }

    private static T[] Read<T>(Action<OneToOneMap<T>> configure, params object?[] cells)
    {
        var rows = new object?[cells.Length + 1][];
        rows[0] = ["Value"];
        for (int i = 0; i < cells.Length; i++)
        {
            rows[i + 1] = [cells[i]];
        }

        using var importer = XlsxBuilder.Create(rows);
        importer.Configuration.RegisterClassMap<Row<T>>(c => configure(c.Map(r => r.Value)));
        return [.. importer.ReadSheet().ReadRows<Row<T>>().Select(r => r.Value)];
    }

    public class Row<T>
    {
        public T Value { get; set; } = default!;
    }
}

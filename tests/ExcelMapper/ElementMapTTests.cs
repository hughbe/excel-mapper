using System.Globalization;

namespace ExcelMapper.Tests;

public class ElementMapTTests
{
    [Fact]
    public void Pipeline_CollectionMap_IsAnElementMap()
    {
        var collectionMap = new ExcelClassMap<Row>().Map(r => r.Dates);
        var dictionaryMap = new ExcelClassMap<Row>().Map(r => r.Values);
        Assert.IsType<ElementMap<DateTime>>(Helpers.ElementsOf(collectionMap));
        Assert.IsType<ElementMap<int?>>(Helpers.ValuesOf(dictionaryMap));
    }

    [Fact]
    public void TypeMethods_ElementMap_ReturnTheMap()
    {
        var map = Helpers.ElementsOf(new ExcelClassMap<Row>().Map(r => r.Dates));
        Assert.Same(map, map.WithFormats("yyyy-MM-dd"));
        Assert.Same(map, map.WithFormatProvider(CultureInfo.InvariantCulture));

        var nullableMap = Helpers.ValuesOf(new ExcelClassMap<Row>().Map(r => r.Values));
        Assert.Same(nullableMap, nullableMap.WithConverter(new NumberConverter()));
        Assert.Same(nullableMap, nullableMap.WithEmptyFallback(new ZeroFallback()));
        Assert.Same(nullableMap, nullableMap.WithInvalidFallback(new ZeroFallback()));
    }

    [Fact]
    public void ReadRows_ElementMapMethods_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Dates,Values:a,Values:b\n2024/01/02,1,\n");
        importer.Configuration.RegisterClassMap<Row>(c =>
        {
            c.Map(r => r.Dates)
                .WithColumnNames("Dates")
                .WithElementMap(e => e.WithTrim().WithFormats("yyyy/MM/dd"));
            c.Map(r => r.Values)
                .WithColumnNames("Values:a", "Values:b")
                .WithValueMap(v => v.WithConverter(new NumberConverter()).WithEmptyFallback(new ZeroFallback()));
        });

        var row = Assert.Single(importer.ReadSheet().ReadRows<Row>());
        Assert.Equal([new DateTime(2024, 1, 2)], row.Dates);
        Assert.Equal(new Dictionary<string, int?> { { "Values:a", 10 }, { "Values:b", 0 } }, row.Values);
    }

    public class Row
    {
        public DateTime[] Dates { get; set; } = default!;
        public Dictionary<string, int?> Values { get; set; } = default!;
    }

    // Multiplies by ten, to show the converter is used.
    private sealed class NumberConverter : ICellConverter<int>
    {
        public CellConversionResult<int> Convert(CellValue value, CellContext context)
            => int.TryParse(value.Text, out var result) ? CellConversionResult<int>.Success(result * 10) : CellConversionResult<int>.Invalid();
    }

    private sealed class ZeroFallback : ICellFallback<int>
    {
        public int Fallback(CellValue value, CellContext context, Exception? exception) => 0;
    }
}

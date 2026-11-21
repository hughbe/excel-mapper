using System.Linq;

namespace ExcelMapper.Tests;

/// <summary>
/// The fluent fallback methods: throwing fallbacks, value fallbacks for elements, and fallbacks for
/// the underlying type of nullable members.
/// </summary>
public class MapFallbackMethodsTests
{
    [Fact]
    public void ReadRows_WithThrowingFallback_ThrowsForEmptyAndInvalidCells()
    {
        // A nullable member is null for empty cells by default.
        using var importer = Helpers.GetCsvImporter("Value\n1\n\nabc\n");
        importer.Configuration.RegisterClassMap<NullableRow>(c => c.Map(r => r.Value).WithThrowingFallback());
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();
        Assert.Equal(1, sheet.ReadRow<NullableRow>().Value);
        Assert.Equal("A3", Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableRow>()).CellAddress);
        Assert.Equal("A4", Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableRow>()).CellAddress);
    }

    [Fact]
    public void ReadRows_WithThrowingEmptyFallback_ThrowsForEmptyCells()
    {
        using var importer = Helpers.GetCsvImporter("Value,Other\nabc,1\n,2\n");
        importer.Configuration.RegisterClassMap<NullableRow>(c => c.Map(r => r.Value).WithInvalidFallback(-1).WithThrowingEmptyFallback());
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();
        Assert.Equal(-1, sheet.ReadRow<NullableRow>().Value);
        Assert.Equal("A3", Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableRow>()).CellAddress);
    }

    [Fact]
    public void ReadRows_WithThrowingInvalidFallback_ReplacesEarlierInvalidFallback()
    {
        using var importer = Helpers.GetCsvImporter("Value\n\nabc\n");
        importer.Configuration.RegisterClassMap<NullableRow>(c => c.Map(r => r.Value).WithValueFallback(-1).WithThrowingInvalidFallback());
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();
        Assert.Equal(-1, sheet.ReadRow<NullableRow>().Value);
        Assert.Equal("A3", Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableRow>()).CellAddress);
    }

    [Fact]
    public void ReadRows_ElementsWithValueFallback_UsedForEmptyAndInvalidElements()
    {
        using var importer = Helpers.GetCsvImporter("Values\n\"1,,abc,4\"\n");
        importer.Configuration.RegisterClassMap<ElementsRow>(c => c.Map(r => r.Values).WithElementMap(e => e.WithValueFallback(0)));
        Assert.Equal([1, 0, 0, 4], importer.ReadSheet().ReadRows<ElementsRow>().Single().Values);
    }

    [Fact]
    public void ReadRows_NullableMemberWithUnderlyingTypeFallbacks_UsesFallbacks()
    {
        using var importer = Helpers.GetCsvImporter("Value\n1\n\nabc\n");
        importer.Configuration.RegisterClassMap<NullableRow>(c => c.Map(r => r.Value)
            .WithEmptyFallback(new FixedFallback(7))
            .WithInvalidFallback(new FixedFallback(-1)));
        Assert.Equal([1, 7, -1], importer.ReadSheet().ReadRows<NullableRow>().Select(r => r.Value));
    }

    [Fact]
    public void ReadRows_FormatsAfterConverter_ConvertsValuesTheConverterDoesNotHandle()
    {
        // The converter replaces the built-in conversion, and WithFormats adds it back after it.
        using var importer = Helpers.GetCsvImporter("Value\ntoday\n20240304\n");
        importer.Configuration.RegisterClassMap<DateRow>(c => c.Map(r => r.Value).WithConverter(new TodayConverter()).WithFormats("yyyyMMdd"));
        Assert.Equal([new DateTime(2000, 1, 1), new DateTime(2024, 3, 4)], importer.ReadSheet().ReadRows<DateRow>().Select(r => r.Value));
    }

    private sealed class FixedFallback(int value) : ICellFallback<int>
    {
        public int Fallback(CellValue cell, CellContext context, Exception? exception) => value;
    }

    private sealed class TodayConverter : ICellConverter<DateTime>
    {
        public CellConversionResult<DateTime> Convert(CellValue value, CellContext context)
            => value.Text == "today" ? CellConversionResult<DateTime>.Success(new DateTime(2000, 1, 1)) : CellConversionResult<DateTime>.NotHandled;
    }

    public class NullableRow
    {
        public int? Value { get; set; }
    }

    public class ElementsRow
    {
        public int[] Values { get; set; } = [];
    }

    public class DateRow
    {
        public DateTime Value { get; set; }
    }
}

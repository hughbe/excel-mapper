using System.Linq;

namespace ExcelMapper.Tests;

public class ExcelRegionTests
{
    [Fact]
    public void WithHeadingAt_Invoke_ReturnsNewRegion()
    {
        using var importer = Helpers.GetCsvImporter("A\n1\n");
        var sheet = importer.ReadSheet();
        var region = sheet.Region("A1:A2");
        var result = region.WithHeadingAt(1);
        Assert.NotSame(region, result);
        Assert.Same(sheet, result.Sheet);
        Assert.Equal(region.Range, result.Range);
    }

    [Fact]
    public void WithHeadingAt_NegativeRowIndex_ThrowsArgumentOutOfRangeException()
    {
        using var importer = Helpers.GetCsvImporter("A\n1\n");
        var region = importer.ReadSheet().Region();
        Assert.Throws<ArgumentOutOfRangeException>("rowIndex", () => region.WithHeadingAt(-1));
    }

    [Fact]
    public void WithoutHeading_Invoke_ReturnsNewRegion()
    {
        using var importer = Helpers.GetCsvImporter("A\n1\n");
        var region = importer.ReadSheet().Region();
        var result = region.WithoutHeading();
        Assert.NotSame(region, result);
        Assert.Same(region.Sheet, result.Sheet);
        Assert.Equal(region.Range, result.Range);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(50)]
    public void WithHeadingSearch_Invoke_ReturnsNewRegion(int maxRows)
    {
        using var importer = Helpers.GetCsvImporter("A\n1\n");
        var region = importer.ReadSheet().Region();
        Assert.NotSame(region, region.WithHeadingSearch());
        var result = region.WithHeadingSearch(maxRows);
        Assert.NotSame(region, result);
        Assert.Equal(region.Range, result.Range);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    public void WithHeadingSearch_InvalidMaxRows_ThrowsArgumentOutOfRangeException(int maxRows)
    {
        using var importer = Helpers.GetCsvImporter("A\n1\n");
        var region = importer.ReadSheet().Region();
        Assert.Throws<ArgumentOutOfRangeException>("maxRows", () => region.WithHeadingSearch(maxRows));
    }

    [Fact]
    public void ReadRows_Closed_ThrowsExcelMappingException()
    {
        var importer = Helpers.GetCsvImporter("A\n1\n");
        var region = importer.ReadSheet().Region();
        importer.Dispose();
        var exception = Assert.Throws<ExcelMappingException>(() => region.ReadRows<StringValue>().GetEnumerator().MoveNext());
        Assert.Equal("The underlying reader is closed.", exception.Message);
    }

    [Fact]
    public void ReadRows_IsLazy()
    {
        using var importer = Helpers.GetCsvImporter("A\n1\n");
        var region = importer.ReadSheet().Region();

        // Nothing is read until the result is enumerated.
        var rows = region.ReadRows<StringValue>();
        var results = region.ReadRowResults<StringValue>();
        Assert.Equal(-1, region.Sheet.CurrentRowIndex);
        Assert.NotNull(rows);
        Assert.NotNull(results);
    }

    private class StringValue
    {
        [ExcelColumnIndex(0)]
        public string? Value { get; set; }
    }

    private static ExcelClassMap<RegionMapRow> CreateGivenMap()
    {
        var classMap = new ExcelClassMap<RegionMapRow>();
        classMap.Map(r => r.Name).WithColumnName("Other");
        return classMap;
    }

    public class RegionMapRow
    {
        public string? Name { get; set; }
    }

    [Fact]
    public void ReadRows_ClassMap_UsesClassMap()
    {
        using var importer = Helpers.GetCsvImporter("Title,,\n,Name,Other\n,a,x\n");
        var region = importer.ReadSheet().Region("B2:C3");
        Assert.Equal(["x"], region.ReadRows(CreateGivenMap()).Select(r => r.Name));
        Assert.Equal(["x"], region.ReadRowResults(CreateGivenMap()).Select(r => r.Value!.Name));
        region.ValidateHeading(CreateGivenMap());
        Assert.True(region.TryValidateHeading(CreateGivenMap(), out _));
    }

    [Fact]
    public void ReadRows_ClassMapHeadingSearch_FindsHeading()
    {
        using var importer = Helpers.GetCsvImporter("Title,\nName,Other\na,x\n");
        var sheet = importer.ReadSheet();
        var region = sheet.Region().WithHeadingSearch();
        Assert.Equal(["x"], region.ReadRows(CreateGivenMap()).Select(r => r.Name));
        Assert.True(region.TryValidateHeading(CreateGivenMap(), out _));

        var missing = sheet.Region("A1:A1").WithHeadingSearch(1);
        Assert.False(missing.TryValidateHeading(CreateGivenMap(), out var problems));
        Assert.Single(problems);
    }

    [Fact]
    public void ReadRows_NullClassMap_ThrowsArgumentNullException()
    {
        using var importer = Helpers.GetCsvImporter("Name\na\n");
        var region = importer.ReadSheet().Region();
        Assert.Throws<ArgumentNullException>("classMap", () => region.ReadRows<RegionMapRow>(null!));
        Assert.Throws<ArgumentNullException>("classMap", () => region.ReadRowResults<RegionMapRow>(null!));
        Assert.Throws<ArgumentNullException>("classMap", () => region.ValidateHeading<RegionMapRow>(null!));
        Assert.Throws<ArgumentNullException>("classMap", () => region.TryValidateHeading<RegionMapRow>(null!, out _));
    }
}

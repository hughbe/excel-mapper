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
    public void ReadRows_Closed_ThrowsObjectDisposedException()
    {
        var importer = Helpers.GetCsvImporter("A\n1\n");
        var region = importer.ReadSheet().Region();
        importer.Dispose();
        var exception = Assert.Throws<ObjectDisposedException>(() => region.ReadRows<StringValue>().GetEnumerator().MoveNext());
        Assert.Equal(nameof(ExcelImporter), exception.ObjectName);
    }

    [Fact]
    public void ReadRows_DoesNotMoveSheet()
    {
        using var importer = Helpers.GetCsvImporter("A\n1\n");
        var region = importer.ReadSheet().Region();

        // A region reads through its own view: the sheet's position doesn't change.
        var rows = region.ReadRows<StringValue>();
        var results = region.ReadRowResults<StringValue>();
        Assert.Equal(-1, region.Sheet.CurrentRowIndex);
        Assert.NotNull(rows);
        Assert.NotNull(results);
    }

    [Fact]
    public void ReadMethods_MissingColumn_ThrowWhenCalled()
    {
        using var importer = Helpers.GetCsvImporter("A\n1\n");
        var region = importer.ReadSheet().Region();

        // As for a sheet, problems with the heading are reported when the method is called,
        // rather than when the result is enumerated.
        Assert.Throws<ExcelMappingException>(() => region.ReadRowResults<RegionMapRow>());
        var exception = Assert.Throws<ExcelMappingException>(() => region.ReadRows<RegionMapRow>(new CountingSink()));
        Assert.StartsWith("The sheet is missing columns needed to map type ExcelRegionTests.RegionMapRow", exception.Message);
    }

    [Fact]
    public void ReadRows_HeadingNotFound_ThrowsWhenCalled()
    {
        using var importer = Helpers.GetCsvImporter("A\n1\n");
        var region = importer.ReadSheet().Region().WithHeadingSearch();
        Assert.Throws<ExcelMappingException>(() => region.ReadRows<RegionMapRow>());
    }

    [Fact]
    public void ReadRows_EnumeratedTwice_ReadsRegionAgain()
    {
        using var importer = Helpers.GetCsvImporter("A\n1\n2\n");
        var rows = importer.ReadSheet().Region().ReadRows<StringValue>();
        Assert.Equal(["1", "2"], rows.Select(r => r.Value));
        Assert.Equal(["1", "2"], rows.Select(r => r.Value));
    }

    private sealed class CountingSink : IExcelErrorSink
    {
        public int Count { get; private set; }

        public void Add(in ExcelMappingError error) => Count++;
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
        Assert.Throws<ArgumentNullException>("classMap", () => region.ReadRows<RegionMapRow>((ExcelClassMap<RegionMapRow>)null!));
        Assert.Throws<ArgumentNullException>("classMap", () => region.ReadRowResults<RegionMapRow>(null!));
        Assert.Throws<ArgumentNullException>("classMap", () => region.ValidateHeading<RegionMapRow>(null!));
        Assert.Throws<ArgumentNullException>("classMap", () => region.TryValidateHeading<RegionMapRow>(null!, out _));
    }

    private sealed class RegionCountingSink : IExcelErrorSink
    {
        public int Count { get; private set; }

        public void Add(in ExcelMappingError error) => Count++;
    }

    public class RegionSinkRow
    {
        public int Value { get; set; }
    }

    [Fact]
    public void ReadRows_Sink_ReturnsValidRowsAndReportsOthers()
    {
        using var importer = Helpers.GetCsvImporter("Title\nValue,Other\n1,a\nx,b\n3,c\n");
        var sink = new RegionCountingSink();
        var region = importer.ReadSheet().Region("A2:B5");

        Assert.Equal([1, 3], region.ReadRows<RegionSinkRow>(sink).Select(r => r.Value));
        Assert.Equal(1, sink.Count);

        var map = new ExcelClassMap<RegionSinkRow>();
        map.Map(p => p.Value);
        Assert.Equal([1, 3], region.ReadRows(map, sink).Select(r => r.Value));
        Assert.Equal(2, sink.Count);
    }

    [Fact]
    public void ReadRows_NullSink_ThrowsArgumentNullException()
    {
        using var importer = Helpers.GetCsvImporter("Value\n1\n");
        var region = importer.ReadSheet().Region();
        Assert.Throws<ArgumentNullException>("errors", () => region.ReadRows<RegionSinkRow>((IExcelErrorSink)null!));
        Assert.Throws<ArgumentNullException>("errors", () => region.ReadRows(new ExcelClassMap<RegionSinkRow>(), null!));
    }
}

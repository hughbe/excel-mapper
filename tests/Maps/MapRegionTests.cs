using System.Linq;

namespace ExcelMapper.Tests;

public class MapRegionTests
{
    // Two tables side by side, both with a "Name" column, and a title above them.
    private const string SideBySide =
        "Sales,,,Costs,\n" +
        "Name,Amount,,Name,Cost\n" +
        "a,1,,x,10\n" +
        "b,2,,y,20\n" +
        ",,,z,30\n";

    [Fact]
    public void ReadRows_TwoTablesSideBySide_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter(SideBySide);
        var sheet = importer.ReadSheet();
        var sales = sheet.Region("A2:B4").ReadRows<Sale>().ToArray();
        var costs = sheet.Region("D2:E5").ReadRows<CostRow>().ToArray();

        Assert.Equal(["a", "b"], sales.Select(s => s.Name));
        Assert.Equal([1, 2], sales.Select(s => s.Amount));
        Assert.Equal(["x", "y", "z"], costs.Select(c => c.Name));
        Assert.Equal([10, 20, 30], costs.Select(c => c.Cost));
    }

    [Fact]
    public void ReadRows_TwoTablesStacked_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Name,Amount\na,1\nb,2\n,\nName,Cost\nx,10\n");
        var sheet = importer.ReadSheet();
        Assert.Equal(["a", "b"], sheet.Region("1:3").ReadRows<Sale>().Select(s => s.Name));
        Assert.Equal([10], sheet.Region("5:6").ReadRows<CostRow>().Select(c => c.Cost));
    }

    [Fact]
    public void ReadRows_InterleavedRegions_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter(SideBySide);
        var sheet = importer.ReadSheet();
        var pairs = sheet.Region("A2:B4").ReadRows<Sale>()
            .Zip(sheet.Region("D2:E5").ReadRows<CostRow>(), (s, c) => $"{s.Name}{c.Name}")
            .ToArray();
        Assert.Equal(["ax", "by"], pairs);
    }

    [Fact]
    public void ReadRows_ReadTwice_ReadsFromTheTopEachTime()
    {
        using var importer = Helpers.GetCsvImporter(SideBySide);
        var region = importer.ReadSheet().Region("A2:B4");
        var rows = region.ReadRows<Sale>();
        Assert.Equal(["a", "b"], rows.Select(s => s.Name));
        Assert.Equal(["a", "b"], rows.Select(s => s.Name));
        Assert.Equal(["a", "b"], region.ReadRows<Sale>().Select(s => s.Name));
    }

    [Fact]
    public void ReadRows_DoesNotChangeSheet()
    {
        using var importer = Helpers.GetCsvImporter("Name,Amount\na,1\nb,2\nc,3\n");
        var sheet = importer.ReadSheet();
        Assert.Equal("a", sheet.ReadRow<Sale>().Name);

        Assert.Equal(["b", "c"], sheet.Region("2:4").WithHeadingAt(0).ReadRows<Sale>().Select(s => s.Name).Skip(1));
        Assert.Equal(1, sheet.CurrentRowIndex);
        Assert.Equal(default, sheet.DataRange);
        Assert.Equal("b", sheet.ReadRow<Sale>().Name);
    }

    [Fact]
    public void ReadRows_HeadingAboveRegion_ReadsEveryRowOfRegion()
    {
        // For example, reading rows 3 to 4 of a sheet with its heading in row 1.
        using var importer = Helpers.GetCsvImporter("Name,Amount\na,1\nb,2\nc,3\nd,4\n");
        var rows = importer.ReadSheet().Region(new ExcelRange(2..4, Range.All)).WithHeadingAt(0).ReadRows<Sale>().ToArray();
        Assert.Equal(["b", "c"], rows.Select(s => s.Name));
        Assert.Equal([2, 3], rows.Select(s => s.Amount));
    }

    [Fact]
    public void ReadRows_HeadingInsideRegion_ReadsRowsAfterIt()
    {
        using var importer = Helpers.GetCsvImporter("Title,\nNotes,\nName,Amount\na,1\nb,2\n");
        var rows = importer.ReadSheet().Region().WithHeadingAt(2).ReadRows<Sale>().ToArray();
        Assert.Equal(["a", "b"], rows.Select(s => s.Name));
    }

    [Fact]
    public void ReadRows_HeadingBelowRegion_ReturnsNoRows()
    {
        using var importer = Helpers.GetCsvImporter("Name,Amount\na,1\nb,2\n");
        Assert.Empty(importer.ReadSheet().Region("1:2").WithHeadingAt(2).ReadRows<Sale>());
    }

    [Fact]
    public void ReadRows_WithoutHeading_ReadsEveryRow()
    {
        using var importer = Helpers.GetCsvImporter("a,1\nb,2\n");
        var rows = importer.ReadSheet().Region().WithoutHeading().ReadRows<Positional>().ToArray();
        Assert.Equal(["a", "b"], rows.Select(r => r.Name));
        Assert.Equal([1, 2], rows.Select(r => r.Amount));
    }

    [Fact]
    public void ReadRows_WithoutHeadingSubRegion_ColumnPositionsCountFromRegion()
    {
        // Column 0 (letter A) is the region's first column, so one type maps any table of that shape.
        using var importer = Helpers.GetCsvImporter("x,a,1\nx,b,2\nx,c,3\n");
        var rows = importer.ReadSheet().Region("B2:C3").WithoutHeading().ReadRows<Positional>().ToArray();
        Assert.Equal(["b", "c"], rows.Select(r => r.Name));
        Assert.Equal([2, 3], rows.Select(r => r.Amount));
    }

    [Fact]
    public void ReadRows_TwoTablesWithoutHeadings_ShareOneType()
    {
        using var importer = Helpers.GetCsvImporter("a,1,,x,10\nb,2,,y,20\n");
        var sheet = importer.ReadSheet();
        var left = sheet.Region("A1:B2").WithoutHeading().ReadRows<Positional>().ToArray();
        var right = sheet.Region("D1:E2").WithoutHeading().ReadRows<Positional>().ToArray();
        Assert.Equal(["a", "b"], left.Select(r => r.Name));
        Assert.Equal(["x", "y"], right.Select(r => r.Name));
        Assert.Equal([10, 20], right.Select(r => r.Amount));
    }

    [Fact]
    public void ReadRows_ColumnLettersCountFromRegion()
    {
        using var importer = Helpers.GetCsvImporter("x,Name,Amount\nx,a,1\n");
        var row = importer.ReadSheet().Region("B1:C2").ReadRows<ByLetter>().Single();
        Assert.Equal("a", row.Name);
        Assert.Equal(1, row.Amount);
    }

    public class ByLetter
    {
        [ExcelColumnLetter("A")]
        public string? Name { get; set; }

        [ExcelColumnLetter("B")]
        public int Amount { get; set; }
    }

    [Fact]
    public void ReadRows_ColumnIndicesCountFromRegion()
    {
        using var importer = Helpers.GetCsvImporter("x,1,2,3\n");
        var row = importer.ReadSheet().Region("B1:D1").WithoutHeading().ReadRows<Indices>().Single();
        Assert.Equal([1, 3], row.Values);
    }

    public class Indices
    {
        [ExcelColumnIndices(0, 2)]
        public int[] Values { get; set; } = [];
    }

    [Fact]
    public void ReadRows_ColumnPositionPastRegion_IsMissing()
    {
        // The worksheet has column D, but the region stops at C.
        using var importer = Helpers.GetCsvImporter("a,1,x,y\n");
        var region = importer.ReadSheet().Region("B1:C1").WithoutHeading();
        Assert.Throws<ExcelMappingException>(() => region.ReadRows<ThirdColumn>().ToArray());
        Assert.False(region.TryValidateHeading<ThirdColumn>(out var problems));
        Assert.Contains("column D (index 2)", Assert.Single(problems));
    }

    public class ThirdColumn
    {
        [ExcelColumnIndex(2)]
        public string? Value { get; set; }
    }

    [Fact]
    public void ReadRows_ErrorAddressForColumnPosition_IsInWorksheet()
    {
        using var importer = Helpers.GetCsvImporter("x,a,bad\n");
        var region = importer.ReadSheet().Region("B1:C1").WithoutHeading();
        var exception = Assert.Throws<ExcelMappingException>(() => region.ReadRows<Positional>().ToArray());
        Assert.Equal("C1", exception.CellAddress);
    }

    [Fact]
    public void ReadRows_SheetColumnPositionsUnchanged()
    {
        // Positions on the sheet itself (and with DataRange) are still worksheet positions.
        using var importer = Helpers.GetCsvImporter("a,1\nb,2\n");
        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;
        Assert.Equal(["a", "b"], sheet.ReadRows<Positional>().Select(r => r.Name));
    }

    [Fact]
    public void ReadRows_BlankRowInRegion_MappedUnlessSkipped()
    {
        // Row 5 is blank in columns A:B (its values are in D:E).
        using var importer = Helpers.GetCsvImporter(SideBySide);
        var sheet = importer.ReadSheet();
        var exception = Assert.Throws<ExcelMappingException>(() => sheet.Region(new ExcelRange(1.., 0..2)).ReadRows<Sale>().ToArray());
        Assert.Equal("B5", exception.CellAddress);

        importer.Configuration.SkipBlankLines = true;
        Assert.Equal(["a", "b"], sheet.Region(new ExcelRange(1.., 0..2)).ReadRows<Sale>().Select(s => s.Name));
    }

    [Fact]
    public void ReadRows_WithHeadingSearch_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter(SideBySide);
        var sheet = importer.ReadSheet();
        Assert.Equal(["a", "b"], sheet.Region("A1:B4").WithHeadingSearch().ReadRows<Sale>().Select(s => s.Name));
        Assert.Equal(["x", "y", "z"], sheet.Region("D:E").WithHeadingSearch(5).ReadRows<CostRow>().Select(c => c.Name));
    }

    [Fact]
    public void ReadRows_WithHeadingSearchNotFound_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetCsvImporter(SideBySide);
        var region = importer.ReadSheet().Region("A:B").WithHeadingSearch(1);
        var exception = Assert.Throws<ExcelMappingException>(() => region.ReadRows<Sale>().ToArray());
        Assert.Contains("Could not find the heading", exception.Message);
    }

    [Fact]
    public void ReadRows_RegionFromEnd_ExcludesTotalsRow()
    {
        using var importer = Helpers.GetCsvImporter("Name,Amount\na,1\nb,2\nTotal,3\n");
        var rows = importer.ReadSheet().Region(new ExcelRange(..^1, Range.All)).ReadRows<Sale>().ToArray();
        Assert.Equal(["a", "b"], rows.Select(s => s.Name));
    }

    [Fact]
    public void ReadRows_Xlsx_ReturnsExpected()
    {
        using var importer = XlsxBuilder.Create(["Report", null, null], [null, "Name", "Amount"], [null, "a", 1.0], [null, "b", 2.0]);
        var rows = importer.ReadSheet().Region("B2:C4").ReadRows<Sale>().ToArray();
        Assert.Equal(["a", "b"], rows.Select(s => s.Name));
        Assert.Equal([1, 2], rows.Select(s => s.Amount));
    }

    [Fact]
    public void ReadRows_ClassMap_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter(SideBySide);
        importer.Configuration.RegisterClassMap<Order>(c =>
        {
            c.Map(o => o.Id).WithColumnName("Name");
            c.Map(o => o.Total).WithColumnName("Cost");
        });
        var rows = importer.ReadSheet().Region("D2:E5").ReadRows<Order>().ToArray();
        Assert.Equal(["x", "y", "z"], rows.Select(o => o.Id));
    }

    [Fact]
    public void ReadRows_ErrorAddress_IsInWorksheet()
    {
        using var importer = Helpers.GetCsvImporter("Title,,,\n,,Name,Cost\n,,x,bad\n");
        var region = importer.ReadSheet().Region("C2:D3");
        var exception = Assert.Throws<ExcelMappingException>(() => region.ReadRows<CostRow>().ToArray());
        Assert.Equal("D3", exception.CellAddress);
    }

    [Fact]
    public void ReadRowResults_CollectsErrors()
    {
        using var importer = Helpers.GetCsvImporter("Name,Cost\nx,1\ny,bad\nz,3\n");
        var results = importer.ReadSheet().Region().ReadRowResults<CostRow>().ToArray();
        Assert.Equal(3, results.Length);
        Assert.Equal("B3", Assert.Single(results[1].Errors).CellAddress);
        Assert.Equal(3, results[2].Value!.Cost);
    }

    [Fact]
    public void ValidateHeading_MissingColumn_ThrowsAndReports()
    {
        using var importer = Helpers.GetCsvImporter(SideBySide);
        var sheet = importer.ReadSheet();
        sheet.Region("A2:B4").ValidateHeading<Sale>();
        Assert.True(sheet.Region("A2:B4").TryValidateHeading<Sale>(out var problems));
        Assert.Empty(problems);

        Assert.False(sheet.Region("A2:B4").TryValidateHeading<CostRow>(out problems));
        Assert.Contains("\"Cost\"", Assert.Single(problems));
        Assert.Throws<ExcelMappingException>(() => sheet.Region("A2:B4").ValidateHeading<CostRow>());
    }

    [Fact]
    public void ValidateHeading_WithHeadingSearch_ReportsWhenNotFound()
    {
        using var importer = Helpers.GetCsvImporter(SideBySide);
        var sheet = importer.ReadSheet();
        sheet.Region("D:E").WithHeadingSearch().ValidateHeading<CostRow>();

        Assert.False(sheet.Region("A:B").WithHeadingSearch().TryValidateHeading<CostRow>(out var problems));
        Assert.Contains("Could not find the heading", Assert.Single(problems));
        Assert.Throws<ExcelMappingException>(() => sheet.Region("A:B").WithHeadingSearch().ValidateHeading<CostRow>());
    }

    public class Sale
    {
        public string? Name { get; set; }
        public int Amount { get; set; }
    }

    public class CostRow
    {
        public string? Name { get; set; }
        public int Cost { get; set; }
    }

    public class Order
    {
        public string? Id { get; set; }
        public int Total { get; set; }
    }

    public class Positional
    {
        [ExcelColumnIndex(0)]
        public string? Name { get; set; }

        [ExcelColumnIndex(1)]
        public int Amount { get; set; }
    }
}

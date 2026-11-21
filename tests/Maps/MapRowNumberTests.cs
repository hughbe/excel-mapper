using System.Linq;
using System.Text;

namespace ExcelMapper.Tests;

public class MapRowNumberTests
{
    [Fact]
    public void ReadRows_AutoMappedRowNumber_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Name\na\nb\nc\n");
        var rows = importer.ReadSheet().ReadRows<RowNumberRow>().ToArray();
        Assert.Equal(["a", "b", "c"], rows.Select(r => r.Name));
        Assert.Equal([2, 3, 4], rows.Select(r => r.Row));
    }

    public class RowNumberRow
    {
        [ExcelRowNumber]
        public int Row { get; set; }

        public string? Name { get; set; }
    }

    [Fact]
    public void ReadRows_AutoMappedRowNumberXlsx_ReturnsExpected()
    {
        using var importer = XlsxBuilder.Create(["Name"], ["a"], ["b"]);
        var rows = importer.ReadSheet().ReadRows<RowNumberRow>().ToArray();
        Assert.Equal([2, 3], rows.Select(r => r.Row));
    }

    [Fact]
    public void ReadRow_AutoMappedRowNumber_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Name\na\nb\n");
        var sheet = importer.ReadSheet();
        Assert.Equal(2, sheet.ReadRow<RowNumberRow>().Row);
        Assert.Equal(3, sheet.ReadRow<RowNumberRow>().Row);
    }

    [Fact]
    public void ReadRows_AutoMappedRowNumberOtherTypes_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Name\na\n");
        var row = importer.ReadSheet().ReadRows<OtherTypesRow>().Single();
        Assert.Equal(2L, row.Long);
        Assert.Equal(2, row.NullableInt);
        Assert.Equal((short)2, row.Short);
        Assert.Equal(2u, row.UInt);
        Assert.Equal(2d, row.Double);
        Assert.Equal(2m, row.Decimal);
        Assert.Equal(2L, row.Field);
    }

    public class OtherTypesRow
    {
        [ExcelRowNumber]
        public long Long { get; set; }

        [ExcelRowNumber]
        public int? NullableInt { get; set; }

        [ExcelRowNumber]
        public short Short { get; set; }

        [ExcelRowNumber]
        public uint UInt { get; set; }

        [ExcelRowNumber]
        public double Double { get; set; }

        [ExcelRowNumber]
        public decimal Decimal { get; set; }

        [ExcelRowNumber]
        public long Field;
    }

    [Fact]
    public void ReadRows_RowNumberWithoutColumn_DoesNotReadColumn()
    {
        // A member named after a column still gets the row number, not the column.
        using var importer = Helpers.GetCsvImporter("Row,Name\n100,a\n");
        var row = importer.ReadSheet().ReadRows<RowNumberRow>().Single();
        Assert.Equal(2, row.Row);
        Assert.Equal("a", row.Name);
    }

    [Fact]
    public void ReadRows_CustomMappedRowNumber_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Name\na\nb\n");
        importer.Configuration.RegisterClassMap<FluentRow>(c =>
        {
            c.MapRowNumber(p => p.Line);
            c.Map(p => p.Name);
        });
        var rows = importer.ReadSheet().ReadRows<FluentRow>().ToArray();
        Assert.Equal([2L, 3L], rows.Select(r => r.Line));
        Assert.Equal(["a", "b"], rows.Select(r => r.Name));
    }

    public class FluentRow
    {
        public long Line { get; set; }
        public string? Name { get; set; }
    }

    [Fact]
    public void ReadRows_CustomMappedNestedRowNumber_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Name\na\nb\n");
        importer.Configuration.RegisterClassMap<NestedRow>(c =>
        {
            c.MapRowNumber(p => p.Source.Row);
            c.Map(p => p.Name);
        });
        var rows = importer.ReadSheet().ReadRows<NestedRow>().ToArray();
        Assert.Equal([2, 3], rows.Select(r => r.Source.Row));
    }

    public class NestedRow
    {
        public SourceInfo Source { get; set; } = new();
        public string? Name { get; set; }
    }

    public class SourceInfo
    {
        public int Row { get; set; }
    }

    [Fact]
    public void ReadRows_RowNumberWithRowsAboveHeading_CountsEveryRow()
    {
        using var importer = Helpers.GetCsvImporter("Report,\n,\nName,Other\na,\nb,\n");
        var sheet = importer.ReadSheet();
        sheet.HeadingIndex = 2;
        var rows = sheet.ReadRows<RowNumberRow>().ToArray();
        Assert.Equal([4, 5], rows.Select(r => r.Row));
    }

    [Fact]
    public void ReadRows_RowNumberSkipBlankLines_CountsBlankLines()
    {
        using var importer = Helpers.GetCsvImporter("Name,Other\na,\n,\nb,\n");
        importer.Configuration.SkipBlankLines = true;
        var rows = importer.ReadSheet().ReadRows<RowNumberRow>().ToArray();
        Assert.Equal(["a", "b"], rows.Select(r => r.Name));
        Assert.Equal([2, 4], rows.Select(r => r.Row));
    }

    [Fact]
    public void ReadRows_RowNumberStartIndex_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Name\na\nb\nc\n");
        var rows = importer.ReadSheet().ReadRows<RowNumberRow>(2, 2).ToArray();
        Assert.Equal(["b", "c"], rows.Select(r => r.Name));
        Assert.Equal([3, 4], rows.Select(r => r.Row));
    }

    [Fact]
    public void ReadRows_RowNumberNoHeading_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("a\nb\n");
        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;
        var rows = sheet.ReadRows<NoHeadingRow>().ToArray();
        Assert.Equal(["a", "b"], rows.Select(r => r.Value));
        Assert.Equal([1, 2], rows.Select(r => r.Row));
    }

    public class NoHeadingRow
    {
        [ExcelRowNumber]
        public int Row { get; set; }

        [ExcelColumnIndex(0)]
        public string? Value { get; set; }
    }

    [Fact]
    public void ReadRows_RowNumberRecord_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Name\na\nb\n");
        var rows = importer.ReadSheet().ReadRows<RowNumberRecord>().ToArray();
        Assert.Equal([new RowNumberRecord(2, "a"), new RowNumberRecord(3, "b")], rows);
    }

    public record RowNumberRecord([property: ExcelRowNumber] int Row, string Name);

    [Fact]
    public void ReadRowResults_RowNumber_MatchesErrorRows()
    {
        using var importer = Helpers.GetCsvImporter("Name,Age\na,1\nb,x\nc,3\n");
        var results = importer.ReadSheet().ReadRowResults<AgeRow>().ToArray();
        Assert.Equal([2, 4], results.Where(r => r.IsValid).Select(r => r.Value!.Row));
        var error = Assert.Single(results.Single(r => !r.IsValid).Errors);
        Assert.Equal("B3", error.CellAddress);
        Assert.Equal(3, error.RowIndex + 1);
    }

    public class AgeRow
    {
        [ExcelRowNumber]
        public int Row { get; set; }

        public string? Name { get; set; }

        public int Age { get; set; }
    }

    [Fact]
    public void ReadRows_RowNumberTooLarge_ThrowsExcelMappingException()
    {
        var csv = new StringBuilder("Name\n");
        for (int i = 0; i < 300; i++)
        {
            csv.Append("a\n");
        }

        using var importer = Helpers.GetCsvImporter(csv.ToString());
        var rows = importer.ReadSheet().ReadRows<ByteRow>();
        var exception = Assert.Throws<ExcelMappingException>(() => rows.ToArray());
        Assert.Equal(255, exception.RowIndex);
        Assert.Equal(nameof(ByteRow.Row), exception.Member!.Name);
    }

    public class ByteRow
    {
        [ExcelRowNumber]
        public byte Row { get; set; }
    }

    [Fact]
    public void ReadRows_RowNumberInvalidType_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetCsvImporter("Name\na\n");
        var sheet = importer.ReadSheet();
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<StringRowNumberRow>().ToArray());
    }

    public class StringRowNumberRow
    {
        [ExcelRowNumber]
        public string? Row { get; set; }
    }

    [Fact]
    public void ValidateHeading_RowNumber_Success()
    {
        using var importer = Helpers.GetCsvImporter("Name\na\n");
        var sheet = importer.ReadSheet();
        sheet.ValidateHeading<RowNumberRow>();
        Assert.Equal(2, sheet.ReadRows<RowNumberRow>().Single().Row);
    }
}

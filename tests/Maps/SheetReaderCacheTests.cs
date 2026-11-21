using System.Linq;

namespace ExcelMapper.Tests;

/// <summary>
/// The reader a member's map finds for a sheet (for example the column a name was found in) is kept
/// by the sheet, so that it's found once per sheet. These check the reader is per sheet.
/// </summary>
public class SheetReaderCacheTests
{
    [Fact]
    public void ReadRows_OneClassMapOnSheetsWithColumnsInDifferentPlaces_ReadsEachSheetsColumn()
    {
        var map = new ExcelClassMap<int>();
        map.Map(v => v).WithColumnName("Value");
        using var first = Helpers.GetCsvImporter("Value,Other\n1,x\n2,y\n");
        using var second = Helpers.GetCsvImporter("Other,Value\nx,3\ny,4\n");
        first.Configuration.RegisterClassMap(map);
        second.Configuration.RegisterClassMap(map);
        var firstSheet = first.ReadSheet();
        var secondSheet = second.ReadSheet();

        Assert.Equal(1, firstSheet.ReadRow<int>());
        Assert.Equal(3, secondSheet.ReadRow<int>());
        Assert.Equal(2, firstSheet.ReadRow<int>());
        Assert.Equal(4, secondSheet.ReadRow<int>());
    }

    [Fact]
    public void ReadRows_IndexerMapsOnSheetsWithColumnsInDifferentPlaces_ReadsEachSheetsColumns()
    {
        var map = new ExcelClassMap<IndexerRow>();
        map.Map(r => r.Values[0]).WithColumnName("A");
        map.Map(r => r.Values[1]).WithColumnName("B");
        using var first = Helpers.GetCsvImporter("A,B\n1,2\n");
        using var second = Helpers.GetCsvImporter("B,A\n3,4\n");
        first.Configuration.RegisterClassMap(map);
        second.Configuration.RegisterClassMap(map);

        Assert.Equal([1, 2], first.ReadSheet().ReadRows<IndexerRow>().Single().Values);
        Assert.Equal([4, 3], second.ReadSheet().ReadRows<IndexerRow>().Single().Values);
    }

    public class IndexerRow
    {
        public int[] Values { get; set; } = new int[2];
    }

}

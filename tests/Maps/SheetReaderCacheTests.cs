using System.Linq;
using ExcelMapper.Abstractions;
using ExcelMapper.Readers;

namespace ExcelMapper.Tests;

/// <summary>
/// The reader a reader factory creates for a sheet (for example the column a name was found in) is
/// kept by the sheet, so that it's found once per sheet. These check the reader is per sheet and per
/// factory instance.
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

    [Fact]
    public void ReadRows_FactoriesThatAreEqual_EachUsesItsOwnReader()
    {
        // Readers are kept per factory instance, not per equal factory.
        using var importer = Helpers.GetCsvImporter("A,B\n1,2\n");
        var map = new ExcelClassMap<Pair>();
        map.Map(p => p.First).ReaderFactory = new AlwaysEqualFactory("A");
        map.Map(p => p.Second).ReaderFactory = new AlwaysEqualFactory("B");
        importer.Configuration.RegisterClassMap(map);
        var row = importer.ReadSheet().ReadRows<Pair>().Single();
        Assert.Equal(1, row.First);
        Assert.Equal(2, row.Second);
    }

    [Fact]
    public void ReadRows_CollectionFactoriesThatAreEqual_EachUsesItsOwnReader()
    {
        using var importer = Helpers.GetCsvImporter("A,B,C\n1,2,3\n");
        var map = new ExcelClassMap<Lists>();
        map.Map(l => l.First).ReaderFactory = new AlwaysEqualCellsFactory("A", "B");
        map.Map(l => l.Second).ReaderFactory = new AlwaysEqualCellsFactory("C");
        importer.Configuration.RegisterClassMap(map);
        var row = importer.ReadSheet().ReadRows<Lists>().Single();
        Assert.Equal([1, 2], row.First);
        Assert.Equal([3], row.Second);
    }

    public class Pair
    {
        public int First { get; set; }
        public int Second { get; set; }
    }

    public class Lists
    {
        public List<int> First { get; set; } = [];
        public List<int> Second { get; set; } = [];
    }

    private sealed class AlwaysEqualFactory(string columnName) : ICellReaderFactory
    {
        private readonly ColumnNameReaderFactory _inner = new(columnName);

        public ICellReader? GetCellReader(ExcelSheet sheet) => _inner.GetCellReader(sheet);

        public override bool Equals(object? obj) => obj is AlwaysEqualFactory;

        public override int GetHashCode() => 0;
    }

    private sealed class AlwaysEqualCellsFactory(params string[] columnNames) : ICellsReaderFactory
    {
        private readonly ColumnNamesReaderFactory _inner = new(columnNames);

        public ICellsReader? GetCellsReader(ExcelSheet sheet) => _inner.GetCellsReader(sheet);

        public override bool Equals(object? obj) => obj is AlwaysEqualCellsFactory;

        public override int GetHashCode() => 0;
    }
}

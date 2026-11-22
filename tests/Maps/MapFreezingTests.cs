using System.Globalization;

namespace ExcelMapper.Tests;

/// <summary>
/// A class map, and every map under it, is frozen the first time it's used with a sheet: changing
/// it afterwards throws, so what's worked out from it (its plans and bindings) never goes stale.
/// </summary>
public class MapFreezingTests
{
    private const string Csv = "Name,Age,Joined,Values,Extra\na,1,2024-01-02,\"1,2\",x\nb,2,2024-01-03,\"3,4\",y\n";

    public class Row
    {
        public string Name { get; set; } = default!;
        public int Age { get; set; }
        public DateTime Joined { get; set; }
        public int[] Values { get; set; } = default!;
        public int RowNumber { get; set; }
    }

    private static (ExcelSheet Sheet, ExcelClassMap<Row> Map, ValueMap<string> Name, ValueMap<int> Age, ValueMap<DateTime> Joined, CollectionMap<int> Values) ReadFirstRow()
    {
        var importer = Helpers.GetCsvImporter(Csv);
        var map = new ExcelClassMap<Row>();
        var name = map.Map(r => r.Name);
        var age = map.Map(r => r.Age);
        var joined = map.Map(r => r.Joined);
        var values = map.Map(r => r.Values);
        importer.Configuration.RegisterClassMap(map);
        var sheet = importer.ReadSheet();
        Assert.Equal("a", sheet.ReadRow<Row>().Name);
        return (sheet, map, name, age, joined, values);
    }

    [Fact]
    public void Map_BeforeReading_Succeeds()
    {
        // Registering a class map doesn't freeze it: only reading with it does.
        using var importer = Helpers.GetCsvImporter(Csv);
        var map = new ExcelClassMap<Row>();
        map.Map(r => r.Name);
        importer.Configuration.RegisterClassMap(map);
        map.Map(r => r.Age).WithInvalidFallback(-1);

        var row = importer.ReadSheet().ReadRow<Row>();
        Assert.Equal("a", row.Name);
        Assert.Equal(1, row.Age);
    }

    [Fact]
    public void ClassMap_ChangedAfterReading_ThrowsInvalidOperationException()
    {
        var (_, map, _, _, _, _) = ReadFirstRow();

        Assert.Throws<InvalidOperationException>(() => map.Map(r => r.Name));
        Assert.Throws<InvalidOperationException>(() => map.MapRowNumber(r => r.RowNumber));
        Assert.Throws<InvalidOperationException>(() => map.Map(r => r.Values[0]));
        Assert.Throws<InvalidOperationException>(() => map.WithClassMap(c => c.Map(r => r.Name)));
    }

    [Fact]
    public void ValueMap_ChangedAfterReading_ThrowsInvalidOperationException()
    {
        var (sheet, _, name, age, joined, _) = ReadFirstRow();

        Assert.Throws<InvalidOperationException>(() => name.WithColumnName("Extra"));
        Assert.Throws<InvalidOperationException>(() => name.MakeOptional());
        Assert.Throws<InvalidOperationException>(() => name.MakePreserveFormatting());
        Assert.Throws<InvalidOperationException>(() => name.WithTrim());
        Assert.Throws<InvalidOperationException>(() => name.WithEmptyFallback("none"));
        Assert.Throws<InvalidOperationException>(() => age.WithInvalidFallback(-1));
        Assert.Throws<InvalidOperationException>(() => age.WithConverter(value => 0));
        Assert.Throws<InvalidOperationException>(() => age.WithMapping(new Dictionary<string, int> { ["b"] = 2 }));
        Assert.Throws<InvalidOperationException>(() => joined.WithFormats("yyyy-MM-dd"));
        Assert.Throws<InvalidOperationException>(() => joined.WithFormatProvider(CultureInfo.InvariantCulture));

        // The map reads the next row as it did the first.
        var row = sheet.ReadRow<Row>();
        Assert.Equal("b", row.Name);
        Assert.Equal(2, row.Age);
    }

    [Fact]
    public void CollectionMap_ChangedAfterReading_ThrowsInvalidOperationException()
    {
        var (sheet, _, _, _, _, values) = ReadFirstRow();

        Assert.Throws<InvalidOperationException>(() => values.WithColumnName("Extra"));
        Assert.Throws<InvalidOperationException>(() => values.MakeOptional());
        Assert.Throws<InvalidOperationException>(() => values.WithElementMap(e => e.WithInvalidFallback(0)));
        Assert.Equal([3, 4], sheet.ReadRow<Row>().Values);
    }

    public class DictionaryRow
    {
        public Dictionary<string, string> Values { get; set; } = default!;
    }

    [Fact]
    public void DictionaryMap_ChangedAfterReading_ThrowsInvalidOperationException()
    {
        using var importer = Helpers.GetCsvImporter("A,B\n1,2\n");
        var map = new ExcelClassMap<DictionaryRow>();
        var values = map.Map(r => r.Values);
        importer.Configuration.RegisterClassMap(map);
        importer.ReadSheet().ReadRow<DictionaryRow>();

        Assert.Throws<InvalidOperationException>(() => values.WithColumnNames("A"));
        Assert.Throws<InvalidOperationException>(() => values.MakeOptional());
        Assert.Throws<InvalidOperationException>(() => values.MakePreserveFormatting());
        Assert.Throws<InvalidOperationException>(() => values.WithValueMap(v => v.WithTrim()));
    }

    public class Parent
    {
        public Child Child { get; set; } = default!;
        public int[] Indexed { get; set; } = default!;
    }

    public class Child
    {
        public string Name { get; set; } = default!;
    }

    [Fact]
    public void NestedAndIndexerMaps_ChangedAfterReading_ThrowsInvalidOperationException()
    {
        using var importer = Helpers.GetCsvImporter("Name,A,B\na,1,2\n");
        var map = new ExcelClassMap<Parent>();
        var childName = map.Map(p => p.Child.Name);
        map.Map(p => p.Indexed[0]).WithColumnName("A");
        importer.Configuration.RegisterClassMap(map);
        importer.ReadSheet().ReadRow<Parent>();

        // Nested class maps and elements mapped one by one are frozen with their parent.
        Assert.Throws<InvalidOperationException>(() => childName.WithColumnName("A"));
        Assert.Throws<InvalidOperationException>(() => map.Map(p => p.Indexed[1]));
    }

    [Fact]
    public void SingleValueMap_ChangedAfterReading_ThrowsInvalidOperationException()
    {
        using var importer = Helpers.GetCsvImporter("Value\n1\n");
        var map = new ExcelClassMap<int>();
        map.Map(v => v);
        importer.Configuration.RegisterClassMap(map);
        Assert.Equal(1, importer.ReadSheet().ReadRow<int>());

        Assert.Throws<InvalidOperationException>(() => map.Map(v => v));
    }

    public class AutoRow
    {
        public string Name { get; set; } = default!;
        public int Age { get; set; }
    }

    [Fact]
    public void AutoMappedClassMap_ChangedAfterReading_ThrowsInvalidOperationException()
    {
        using var importer = Helpers.GetCsvImporter(Csv);
        importer.ReadSheet().ReadRow<AutoRow>();
        Assert.True(importer.Configuration.TryGetClassMap<AutoRow>(out var classMap));

        Assert.Throws<InvalidOperationException>(() => ((ExcelClassMap<AutoRow>)classMap).Map(r => r.Name));
    }

    [Fact]
    public void AutoMappedClassMap_ReplacedAfterReading_UsesNewMap()
    {
        // A frozen auto-mapped class map can still be replaced by registering a new one.
        using var importer = Helpers.GetCsvImporter(Csv);
        var sheet = importer.ReadSheet();
        Assert.Equal(1, sheet.ReadRow<AutoRow>().Age);

        importer.Configuration.RegisterClassMap<AutoRow>(c => c.Map(r => r.Age).WithColumnName("Joined").WithInvalidFallback(-1));
        Assert.Equal(-1, sheet.ReadRow<AutoRow>().Age);
    }

    [Fact]
    public void ValidateHeading_ThenChanged_ThrowsInvalidOperationException()
    {
        using var importer = Helpers.GetCsvImporter(Csv);
        var map = new ExcelClassMap<Row>();
        var name = map.Map(r => r.Name);
        importer.Configuration.RegisterClassMap(map);
        importer.ReadSheet().ValidateHeading<Row>();

        Assert.Throws<InvalidOperationException>(() => name.WithColumnName("Extra"));
    }

    [Fact]
    public void FindHeading_ThenChanged_ThrowsInvalidOperationException()
    {
        using var importer = Helpers.GetCsvImporter("Title\n" + Csv);
        var map = new ExcelClassMap<Row>();
        var name = map.Map(r => r.Name);
        importer.Configuration.RegisterClassMap(map);
        importer.ReadSheet().FindHeading<Row>();

        Assert.Throws<InvalidOperationException>(() => name.WithColumnName("Extra"));
    }
}

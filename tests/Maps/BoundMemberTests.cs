using System.Linq;
using ExcelMapper.Readers;

namespace ExcelMapper.Tests;

/// <summary>
/// Class maps whose members each read one column are bound to a sheet once, and then map each row
/// with typed setters. These check that binding behaves exactly like mapping each member directly.
/// </summary>
public class BoundMemberTests
{
    [Fact]
    public void ReadRows_ReaderFactoryChangedAfterReading_UsesNewColumn()
    {
        using var importer = Helpers.GetCsvImporter("A,B\n1,x\n2,y\n");
        var map = new ExcelClassMap<Row>();
        var nameMap = map.Map(r => r.Name).WithColumnName("A");
        importer.Configuration.RegisterClassMap(map);
        var sheet = importer.ReadSheet();
        Assert.Equal("1", sheet.ReadRow<Row>().Name);

        nameMap.ReaderFactory = new ColumnNameReaderFactory("B");
        Assert.Equal("y", sheet.ReadRow<Row>().Name);
    }

    [Fact]
    public void ReadRows_PropertyAddedAfterReading_MapsIt()
    {
        using var importer = Helpers.GetCsvImporter("Name,Age\na,1\nb,2\n");
        var map = new ExcelClassMap<Row>();
        map.Map(r => r.Name);
        importer.Configuration.RegisterClassMap(map);
        var sheet = importer.ReadSheet();
        var first = sheet.ReadRow<Row>();
        Assert.Equal(0, first.Age);

        map.Map(r => r.Age);
        var second = sheet.ReadRow<Row>();
        Assert.Equal("b", second.Name);
        Assert.Equal(2, second.Age);
    }

    [Fact]
    public void ReadRows_SameClassMapOnTwoSheets_BindsEach()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var primitives = importer.ReadSheet("Primitives");
        var third = importer.ReadSheet("Third Sheet");
        Assert.Equal("a", primitives.ReadRow<StringValueRow>().StringValue);
        Assert.Equal("s1", third.ReadRow<StringValueRow>().StringValue);
        Assert.Equal("b", primitives.ReadRow<StringValueRow>().StringValue);
    }

    [Fact]
    public void ReadRowResults_InvalidValue_KeepsInitialValue()
    {
        using var importer = Helpers.GetCsvImporter("Name,Count\na,x\n");
        var result = importer.ReadSheet().ReadRowResults<WithInitializer>().Single();
        Assert.False(result.IsValid);
        Assert.Equal("B2", Assert.Single(result.Errors).CellAddress);
    }

    [Fact]
    public void ReadRowResults_InvalidValueWithThrowFallbackRecorded_DoesNotSetMember()
    {
        // Errors recorded by a fallback without throwing are handled the same way.
        using var importer = Helpers.GetCsvImporter("Name,Count\na,x\nb,3\n");
        var results = importer.ReadSheet().ReadRowResults<WithInitializer>().ToArray();
        Assert.False(results[0].IsValid);
        Assert.True(results[1].IsValid);
        Assert.Equal(3, results[1].Value!.Count);
        Assert.Equal(["x"], results[1].Value!.Tags);
    }

    public class WithInitializer
    {
        public string? Name { get; set; }

        public int Count { get; set; } = 42;

        [ExcelIgnore]
        public List<string> Tags { get; set; } = ["x"];
    }

    [Fact]
    public void ReadRows_FieldsAndInitOnlyProperties_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Name,Age,Score\na,1,2.5\n");
        var row = importer.ReadSheet().ReadRows<FieldsAndInit>().Single();
        Assert.Equal("a", row.Name);
        Assert.Equal(1, row.Age);
        Assert.Equal(2.5, row.Score);
    }

    public class FieldsAndInit
    {
        public string? Name;
        public int Age;
        public double Score { get; init; }
    }

    [Fact]
    public void ReadRows_Struct_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Name,Age\na,1\nb,2\n");
        var rows = importer.ReadSheet().ReadRows<RowStruct>().ToArray();
        Assert.Equal(["a", "b"], rows.Select(r => r.Name));
        Assert.Equal([1, 2], rows.Select(r => r.Age));
    }

    public struct RowStruct
    {
        public string? Name { get; set; }
        public int Age { get; set; }
    }

    [Fact]
    public void ReadRows_MapTypeDiffersFromMemberType_ConvertsAsBefore()
    {
        using var importer = Helpers.GetCsvImporter("Value\nabc\n");
        importer.Configuration.RegisterClassMap<ObjectRow>(c => c.Map(r => (object?)r.Value));
        Assert.Equal("abc", importer.ReadSheet().ReadRows<ObjectRow>().Single().Value);
    }

    public class ObjectRow
    {
        public string? Value { get; set; }
    }

    [Fact]
    public void ReadRows_OptionalMissingColumn_KeepsInitialValue()
    {
        using var importer = Helpers.GetCsvImporter("Name\na\n");
        var row = importer.ReadSheet().ReadRows<OptionalRow>().Single();
        Assert.Equal("a", row.Name);
        Assert.Equal(7, row.Age);
    }

    public class OptionalRow
    {
        public string? Name { get; set; }

        [ExcelOptional]
        public int Age { get; set; } = 7;
    }

    [Fact]
    public void ReadRows_MissingColumn_ThrowsWithCellDetails()
    {
        using var importer = Helpers.GetCsvImporter("Name\na\n");
        var sheet = importer.ReadSheet();
        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<Row>().ToArray());
        Assert.Equal("Age", exception.Member!.Name);
    }

    [Fact]
    public void ReadRows_Record_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Name,Age,Joined\na,1,\nb,2,\n");
        var rows = importer.ReadSheet().ReadRows<PersonRecord>().ToArray();
        Assert.Equal([new PersonRecord("a", 1, null), new PersonRecord("b", 2, null)], rows);
    }

    public record PersonRecord(string Name, int Age, DateTime? Joined);

    [Fact]
    public void ReadRowResults_RecordInvalidValue_ReportsErrorAndSkipsRow()
    {
        using var importer = Helpers.GetCsvImporter("Name,Age,Joined\na,x,\nb,2,\n");
        var results = importer.ReadSheet().ReadRowResults<PersonRecord>().ToArray();
        Assert.False(results[0].IsValid);
        Assert.Equal("B2", Assert.Single(results[0].Errors).CellAddress);
        Assert.Equal(new PersonRecord("b", 2, null), results[1].Value);
    }

    [Fact]
    public void ReadRows_GetOnlyPropertiesThroughConstructor_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Name,Age,Score\na,1,2.5\n");
        var row = importer.ReadSheet().ReadRows<GetOnly>().Single();
        Assert.Equal("a", row.Name);
        Assert.Equal(1, row.Age);
        Assert.Equal(2.5, row.Score);
    }

    public class GetOnly(string name, int age)
    {
        public string Name { get; } = name;
        public int Age { get; } = age;

        // Not a constructor parameter: set after construction.
        public double Score { get; set; }
    }

    [Fact]
    public void ReadRows_CollectionReaderFactoryChangedAfterReading_UsesNewColumns()
    {
        using var importer = Helpers.GetCsvImporter("A,B,C\n1,2,3\n4,5,6\n");
        var map = new ExcelClassMap<CollectionRow>();
        var valuesMap = map.Map(r => r.Values).WithColumnNames("A", "B");
        importer.Configuration.RegisterClassMap(map);
        var sheet = importer.ReadSheet();
        Assert.Equal([1, 2], sheet.ReadRow<CollectionRow>().Values);

        valuesMap.ReaderFactory = new ColumnNamesReaderFactory("B", "C");
        Assert.Equal([5, 6], sheet.ReadRow<CollectionRow>().Values);
    }

    public class CollectionRow
    {
        public int[] Values { get; set; } = [];
    }

    [Fact]
    public void ReadRows_DictionaryReaderFactoryChangedAfterReading_UsesNewColumns()
    {
        using var importer = Helpers.GetCsvImporter("A,B,C\n1,2,3\n4,5,6\n");
        var map = new ExcelClassMap<DictionaryRow>();
        var valuesMap = map.Map(r => r.Values).WithColumnNames("A");
        importer.Configuration.RegisterClassMap(map);
        var sheet = importer.ReadSheet();
        Assert.Equal(new Dictionary<string, int> { ["A"] = 1 }, sheet.ReadRow<DictionaryRow>().Values);

        valuesMap.ReaderFactory = new ColumnNamesReaderFactory("C");
        Assert.Equal(new Dictionary<string, int> { ["C"] = 6 }, sheet.ReadRow<DictionaryRow>().Values);
    }

    public class DictionaryRow
    {
        public Dictionary<string, int> Values { get; set; } = [];
    }

    [Fact]
    public void ReadRows_NestedObjectsCollectionsAndRowNumbers_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Name,Street,Tags,Extra\na,High St,x;y,e\n");
        var row = importer.ReadSheet().ReadRows<Mixed>().Single();
        Assert.Equal("a", row.Name);
        Assert.Equal("High St", row.Address.Street);
        Assert.Equal(["x", "y"], row.Tags);
        Assert.Equal(2, row.Row);
        Assert.Equal(new Dictionary<string, string> { ["Extra"] = "e" }, row.Other);
    }

    public class Mixed
    {
        public string? Name { get; set; }
        public MixedAddress Address { get; set; } = new();
        [ExcelSeparators(';')]
        public string[] Tags { get; set; } = [];
        [ExcelRowNumber]
        public int Row { get; set; }
        [ExcelExtraColumns]
        public Dictionary<string, string> Other { get; set; } = [];
    }

    public class MixedAddress
    {
        public string? Street { get; set; }
    }

    public class Row
    {
        public string? Name { get; set; }
        public int Age { get; set; }
    }

    public class StringValueRow
    {
        public string? StringValue { get; set; }
    }
}

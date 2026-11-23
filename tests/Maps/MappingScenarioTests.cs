using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelMapper.Tests;

/// <summary>
/// Everyday mapping scenarios: ones that should work, and ones that should fail with a clear
/// <see cref="ExcelMappingException"/> rather than an unrelated exception or wrong data.
/// </summary>
public class MappingScenarioTests
{
    private const string PeopleCsv = "Name,Age\nAlice,30\nBob,25\n";

    // ---------- Scenarios that should work ----------

    public class InitOnlyClass
    {
        public string Name { get; init; } = "";
        public int Age { get; init; }
    }

    [Fact]
    public void ReadRows_InitOnlyProperties_Success()
    {
        using var importer = Helpers.GetCsvImporter(PeopleCsv);
        var row = importer.ReadSheet().ReadRows<InitOnlyClass>().First();
        Assert.Equal("Alice", row.Name);
        Assert.Equal(30, row.Age);
    }

    public class RequiredMembersClass
    {
        public required string Name { get; init; }
        public required int Age { get; init; }
    }

    [Fact]
    public void ReadRows_RequiredMembers_Success()
    {
        using var importer = Helpers.GetCsvImporter(PeopleCsv);
        var row = importer.ReadSheet().ReadRows<RequiredMembersClass>().First();
        Assert.Equal("Alice", row.Name);
        Assert.Equal(30, row.Age);
    }

    [Fact]
    public void ReadRow_HeadingNotRead_ReadsHeadingFirst()
    {
        using var importer = Helpers.GetCsvImporter(PeopleCsv);
        var sheet = importer.ReadSheet();

        var row = sheet.ReadRow<InitOnlyClass>();
        Assert.Equal("Alice", row.Name);
        Assert.NotNull(sheet.Heading);
        Assert.Equal(["Name", "Age"], sheet.Heading.ColumnNames);

        Assert.Equal("Bob", sheet.ReadRow<InitOnlyClass>().Name);
    }

    [Fact]
    public void TryReadRow_HeadingNotRead_ReadsHeadingFirst()
    {
        using var importer = Helpers.GetCsvImporter(PeopleCsv);
        var sheet = importer.ReadSheet();

        Assert.True(sheet.TryReadRow<InitOnlyClass>(out var row));
        Assert.Equal("Alice", row.Name);
        Assert.NotNull(sheet.Heading);
    }

    [Fact]
    public void TryReadRow_EmptySheet_ReturnsFalse()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet("Empty");

        Assert.False(sheet.TryReadRow<InitOnlyClass>(out var row));
        Assert.Null(row);
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<InitOnlyClass>());
    }

    [Fact]
    public void ReadRows_PrimitiveWithHeading_Success()
    {
        using var importer = Helpers.GetCsvImporter("Value\n1\n2\n");
        Assert.Equal([1, 2], importer.ReadSheet().ReadRows<int>());
    }

    [Fact]
    public void ReadRows_PrimitiveWithoutHeading_ReadsFirstColumn()
    {
        using var importer = Helpers.GetCsvImporter("1,x\n2,y\n");
        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;
        Assert.Equal([1, 2], sheet.ReadRows<int>());
    }

    public record IndexedRecord([property: ExcelColumnIndex(1)] string Name, [property: ExcelColumnIndex(0)] int Age);

    [Fact]
    public void ReadRows_ColumnIndexAttributesWithoutHeading_Success()
    {
        using var importer = Helpers.GetCsvImporter("30,Alice\n25,Bob\n");
        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;
        Assert.Equal([new IndexedRecord("Alice", 30), new IndexedRecord("Bob", 25)], sheet.ReadRows<IndexedRecord>());
    }

    public record Person(string Name, int Age);

    public class PersonHolder
    {
        public Person Person { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_AutoMappedNestedRecord_Success()
    {
        using var importer = Helpers.GetCsvImporter(PeopleCsv);
        var row = importer.ReadSheet().ReadRows<PersonHolder>().First();
        Assert.Equal(new Person("Alice", 30), row.Person);
    }

    public record GenericRecord<T>(T Name, int Age);

    [Fact]
    public void ReadRows_GenericRecord_Success()
    {
        using var importer = Helpers.GetCsvImporter(PeopleCsv);
        Assert.Equal(new GenericRecord<string>("Alice", 30), importer.ReadSheet().ReadRows<GenericRecord<string>>().First());
    }

    public class BaseClass
    {
        public string Name { get; set; } = "";
    }

    public class DerivedClass : BaseClass
    {
        public int Age { get; set; }
    }

    [Fact]
    public void ReadRows_InheritedProperties_Success()
    {
        using var importer = Helpers.GetCsvImporter(PeopleCsv);
        var row = importer.ReadSheet().ReadRows<DerivedClass>().First();
        Assert.Equal("Alice", row.Name);
        Assert.Equal(30, row.Age);
    }

    public record BaseRecord(string Name);

    public record DerivedRecord(string Name, int Age) : BaseRecord(Name);

    [Fact]
    public void ReadRows_DerivedRecord_Success()
    {
        using var importer = Helpers.GetCsvImporter(PeopleCsv);
        Assert.Equal(new DerivedRecord("Alice", 30), importer.ReadSheet().ReadRows<DerivedRecord>().First());
    }

    public enum Colour
    {
        Red,
        Green
    }

    public record NullableRecord(Colour? Colour, int? Age);

    [Fact]
    public void ReadRows_NullableRecordParameters_Success()
    {
        using var importer = Helpers.GetCsvImporter("Colour,Age\n,\nGreen,3\n");
        Assert.Equal([new NullableRecord(null, null), new NullableRecord(Colour.Green, 3)], importer.ReadSheet().ReadRows<NullableRecord>());
    }

    public record ListRecord(string Name, List<int> Scores);

    public record ArrayRecord(string Name, int[] Scores);

    [Fact]
    public void ReadRows_RecordWithSplitCollections_Success()
    {
        const string Csv = "Name,Scores\nA,\"1,2\"\n";

        using var listImporter = Helpers.GetCsvImporter(Csv);
        Assert.Equal([1, 2], listImporter.ReadSheet().ReadRows<ListRecord>().Single().Scores);

        using var arrayImporter = Helpers.GetCsvImporter(Csv);
        Assert.Equal([1, 2], arrayImporter.ReadSheet().ReadRows<ArrayRecord>().Single().Scores);
    }

    public record struct MutableRecordStruct(string Name, int Age);

    [Fact]
    public void ReadRows_MutableRecordStruct_Success()
    {
        using var importer = Helpers.GetCsvImporter(PeopleCsv);
        Assert.Equal(new MutableRecordStruct("Alice", 30), importer.ReadSheet().ReadRows<MutableRecordStruct>().First());
    }

    public class FieldsClass
    {
        public string Name = "";
        public int Age;
    }

    [Fact]
    public void ReadRows_PublicFields_Success()
    {
        using var importer = Helpers.GetCsvImporter(PeopleCsv);
        var row = importer.ReadSheet().ReadRows<FieldsClass>().First();
        Assert.Equal("Alice", row.Name);
        Assert.Equal(30, row.Age);
    }

    [Fact]
    public void ReadRows_StringArrayAndDictionaryRows_Success()
    {
        using var arrayImporter = Helpers.GetCsvImporter(PeopleCsv);
        Assert.Equal(["Alice", "30"], arrayImporter.ReadSheet().ReadRows<string[]>().First());

        using var dictionaryImporter = Helpers.GetCsvImporter(PeopleCsv);
        var dictionary = dictionaryImporter.ReadSheet().ReadRows<Dictionary<string, string>>().First();
        Assert.Equal("Alice", dictionary["Name"]);
        Assert.Equal("30", dictionary["Age"]);
    }

    public class CharClass
    {
        public char Letter { get; set; }
    }

    [Fact]
    public void ReadRows_Char_Success()
    {
        using var importer = Helpers.GetCsvImporter("Letter\nx\n");
        Assert.Equal('x', importer.ReadSheet().ReadRows<CharClass>().Single().Letter);
    }

    public class ReadOnlyListClass
    {
        public IReadOnlyList<int> Scores { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_IReadOnlyListFromSplitCell_Success()
    {
        using var importer = Helpers.GetCsvImporter("Scores\n\"1,2\"\n");
        Assert.Equal([1, 2], importer.ReadSheet().ReadRows<ReadOnlyListClass>().Single().Scores);
    }

    public class ColumnsClass
    {
        public string[] Values { get; set; } = default!;
    }

    [Fact]
    public void WithColumnsMatching_Predicate_Success()
    {
        using var importer = Helpers.GetCsvImporter("A1,B,A2\n1,x,2\n");
        importer.Configuration.RegisterClassMap<ColumnsClass>(c => c.Map(p => p.Values).WithColumnsMatching(n => n.StartsWith("A")));
        Assert.Equal(["1", "2"], importer.ReadSheet().ReadRows<ColumnsClass>().Single().Values);
    }

    [Fact]
    public void WithColumnsMatching_Regex_Success()
    {
        using var importer = Helpers.GetCsvImporter("Q1,Total,Q2\n1,x,2\n");
        importer.Configuration.RegisterClassMap<ColumnsClass>(c => c.Map(p => p.Values).WithColumnsMatching(new Regex("^Q[0-9]$")));
        Assert.Equal(["1", "2"], importer.ReadSheet().ReadRows<ColumnsClass>().Single().Values);
    }

    public class NameClass
    {
        public string Name { get; set; } = "";
    }

    [Fact]
    public void WithColumnNameMatching_Regex_Success()
    {
        using var importer = Helpers.GetCsvImporter("Id,Full Name\n1,Alice\n");
        importer.Configuration.RegisterClassMap<NameClass>(c => c.Map(p => p.Name).WithColumnMatching(new Regex("name", RegexOptions.IgnoreCase)));
        Assert.Equal("Alice", importer.ReadSheet().ReadRows<NameClass>().Single().Name);
    }

    [Fact]
    public void WithColumnsMatching_NullPredicateOrRegex_ThrowsArgumentNullException()
    {
        var map = new ExcelClassMap<ColumnsClass>();
        Assert.Throws<ArgumentNullException>("predicate", () => map.Map(p => p.Values).WithColumnsMatching((Func<string, bool>)null!));
        Assert.Throws<ArgumentNullException>("regex", () => map.Map(p => p.Values).WithColumnsMatching((Regex)null!));
        Assert.Throws<ArgumentNullException>("regex", () => new ExcelClassMap<NameClass>().Map(p => p.Name).WithColumnMatching((Regex)null!));
    }

    // ---------- Scenarios that should fail clearly ----------

    public interface IThing
    {
        string Name { get; }
    }

    public abstract class AbstractThing
    {
        public string Name { get; set; } = "";
    }

    public class PrivateConstructorClass
    {
        private PrivateConstructorClass()
        {
        }

        public string Name { get; set; } = "";
    }

    public class UnrelatedConstructorClass
    {
        public UnrelatedConstructorClass(int somethingElse)
        {
        }

        public string Name { get; set; } = "";
    }

    [Fact]
    public void ReadRows_TypesThatCannotBeCreated_ThrowsExcelMappingException()
    {
        AssertMappingFails<IThing>();
        AssertMappingFails<AbstractThing>();
        AssertMappingFails<PrivateConstructorClass>();
        AssertMappingFails<UnrelatedConstructorClass>();
    }

    public class MissingColumnClass
    {
        public string Missing { get; set; } = "";
    }

    [Fact]
    public void ReadRows_MissingColumn_ThrowsExcelMappingException()
    {
        var ex = AssertMappingFails<MissingColumnClass>();
        Assert.Contains("\"Missing\"", ex.Message);
    }

    public class InvalidIntClass
    {
        public int Name { get; set; }
    }

    public record InvalidIntRecord(int Name);

    [Fact]
    public void ReadRows_InvalidValue_ThrowsExcelMappingExceptionWithCell()
    {
        var classEx = AssertMappingFails<InvalidIntClass>();
        Assert.Equal("A2", classEx.CellAddress);

        var recordEx = AssertMappingFails<InvalidIntRecord>();
        Assert.Equal("A2", recordEx.CellAddress);
    }

    public record OverflowRecord(byte Age);

    [Fact]
    public void ReadRows_ValueOutOfRange_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetCsvImporter("Age\n300\n");
        var sheet = importer.ReadSheet();
        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<OverflowRecord>().ToArray());
        Assert.Equal("A2", ex.CellAddress);
    }

    public class DelegateClass
    {
        public Func<int> Name { get; set; } = default!;
    }

    public class TypeClass
    {
        public Type Name { get; set; } = default!;
    }

    public class UnmappableClassPropertyClass
    {
        public ExcelSheet Name { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_UnsupportedMemberTypes_ThrowsExcelMappingException()
    {
        AssertMappingFails<DelegateClass>();
        AssertMappingFails<TypeClass>();
        AssertMappingFails<UnmappableClassPropertyClass>();
    }

    public class CircularClass
    {
        public string Name { get; set; } = "";
        public CircularClass Next { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_CircularReference_ThrowsExcelMappingException()
    {
        var ex = AssertMappingFails<CircularClass>();
        Assert.Equal("Circular reference detected: type MappingScenarioTests.CircularClass references itself through its members. Consider applying the ExcelIgnore attribute to break the cycle.", ex.Message);
    }

    [Fact]
    public void ReadRows_NamedColumnsWithoutHeading_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetCsvImporter(PeopleCsv);
        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;

        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<InitOnlyClass>().ToArray());
        Assert.Contains("There's no heading to find columns by name", ex.Message);
    }

    [Fact]
    public void ReadRow_NoMoreRows_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetCsvImporter("Name\nAlice\n");
        var sheet = importer.ReadSheet();
        sheet.ReadRow<NameClass>();
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NameClass>());
    }

    [Fact]
    public void ReadRows_AfterDispose_ThrowsExcelMappingException()
    {
        var importer = Helpers.GetCsvImporter(PeopleCsv);
        var sheet = importer.ReadSheet();
        importer.Dispose();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<InitOnlyClass>().ToArray());
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<InitOnlyClass>());
        Assert.Throws<ExcelMappingException>(() => sheet.TryReadRow<InitOnlyClass>(out _));
    }

    [Fact]
    public void RegisterClassMap_Twice_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetCsvImporter(PeopleCsv);
        importer.Configuration.RegisterClassMap<NameClass>(c => c.Map(p => p.Name));
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap<NameClass>(c => c.Map(p => p.Name)));
    }

    public class GetOnlyCollectionClass
    {
        public string Name { get; set; } = "";
        public List<int> Scores { get; } = [];
    }

    [Fact]
    public void ReadRows_GetOnlyCollectionNotConstructorParameter_IsNotPopulated()
    {
        // Read-only properties are only mapped through a constructor parameter, so an existing
        // collection instance is left as it is rather than being filled in.
        using var importer = Helpers.GetCsvImporter("Name,Scores\nAlice,\"1,2\"\n");
        var row = importer.ReadSheet().ReadRows<GetOnlyCollectionClass>().Single();
        Assert.Equal("Alice", row.Name);
        Assert.Empty(row.Scores);
    }

    private static ExcelMappingException AssertMappingFails<T>()
    {
        using var importer = Helpers.GetCsvImporter(PeopleCsv);
        var sheet = importer.ReadSheet();
        return Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<T>().ToArray());
    }
}

using System.Linq;

namespace ExcelMapper.Tests;

public class MapConstructorTests
{
    private const string PeopleCsv = "Name,Age\nAlice,30\nBob,25\n";

    public record PositionalRecord(string Name, int Age);

    [Fact]
    public void ReadRows_AutoMappedPositionalRecord_Success()
    {
        using var importer = Helpers.GetCsvImporter(PeopleCsv);
        var rows = importer.ReadSheet().ReadRows<PositionalRecord>().ToArray();
        Assert.Equal([new PositionalRecord("Alice", 30), new PositionalRecord("Bob", 25)], rows);
    }

    public record RecordWithExtraProperty(string Name)
    {
        public int Age { get; set; }
    }

    [Fact]
    public void ReadRows_AutoMappedRecordWithExtraSettableProperty_SetsBoth()
    {
        using var importer = Helpers.GetCsvImporter(PeopleCsv);
        var row = importer.ReadSheet().ReadRows<RecordWithExtraProperty>().First();
        Assert.Equal("Alice", row.Name);
        Assert.Equal(30, row.Age);
    }

    public class GetOnlyClass
    {
        // Parameter names match properties case-insensitively.
        public GetOnlyClass(string NAME, int age)
        {
            Name = NAME;
            Age = age;
        }

        public string Name { get; }

        public int Age { get; }

        // Computed properties that don't match a parameter are ignored.
        public string Description => $"{Name} ({Age})";
    }

    [Fact]
    public void ReadRows_AutoMappedGetOnlyClass_Success()
    {
        using var importer = Helpers.GetCsvImporter(PeopleCsv);
        var row = importer.ReadSheet().ReadRows<GetOnlyClass>().First();
        Assert.Equal("Alice", row.Name);
        Assert.Equal(30, row.Age);
        Assert.Equal("Alice (30)", row.Description);
    }

    public readonly struct GetOnlyStruct
    {
        public GetOnlyStruct(string name, int age)
        {
            Name = name;
            Age = age;
        }

        public string Name { get; }

        public int Age { get; }
    }

    [Fact]
    public void ReadRows_AutoMappedGetOnlyStruct_UsesConstructor()
    {
        using var importer = Helpers.GetCsvImporter(PeopleCsv);
        var row = importer.ReadSheet().ReadRows<GetOnlyStruct>().First();
        Assert.Equal("Alice", row.Name);
        Assert.Equal(30, row.Age);
    }

    public readonly record struct PositionalRecordStruct(string Name, int Age);

    [Fact]
    public void ReadRows_AutoMappedPositionalRecordStruct_Success()
    {
        using var importer = Helpers.GetCsvImporter(PeopleCsv);
        var row = importer.ReadSheet().ReadRows<PositionalRecordStruct>().First();
        Assert.Equal(new PositionalRecordStruct("Alice", 30), row);
    }

    public record OptionalColumnRecord(string Name, [property: ExcelOptional] int Score = 42);

    [Fact]
    public void ReadRows_MissingOptionalColumn_UsesParameterDefault()
    {
        using var importer = Helpers.GetCsvImporter(PeopleCsv);
        var row = importer.ReadSheet().ReadRows<OptionalColumnRecord>().First();
        Assert.Equal(new OptionalColumnRecord("Alice", 42), row);
    }

    public class UnmappedParameterClass
    {
        public UnmappedParameterClass(string name, string suffix = "!")
        {
            Name = name + suffix;
        }

        public string Name { get; }
    }

    [Fact]
    public void ReadRows_ParameterWithoutMember_UsesParameterDefault()
    {
        using var importer = Helpers.GetCsvImporter(PeopleCsv);
        var row = importer.ReadSheet().ReadRows<UnmappedParameterClass>().First();
        Assert.Equal("Alice!", row.Name);
    }

    public class MultipleConstructorsClass
    {
        public MultipleConstructorsClass(string name)
        {
            Name = name;
            Age = -1;
        }

        public MultipleConstructorsClass(string name, int age)
        {
            Name = name;
            Age = age;
        }

        public string Name { get; }

        public int Age { get; }
    }

    [Fact]
    public void ReadRows_MultipleConstructors_UsesConstructorMatchingMostMembers()
    {
        using var importer = Helpers.GetCsvImporter(PeopleCsv);
        var row = importer.ReadSheet().ReadRows<MultipleConstructorsClass>().First();
        Assert.Equal("Alice", row.Name);
        Assert.Equal(30, row.Age);
    }

    [Fact]
    public void ReadRows_CustomMappedGetOnlyClass_Success()
    {
        using var importer = Helpers.GetCsvImporter("Full Name,Years\nAlice,30\n");
        importer.Configuration.RegisterClassMap<GetOnlyClass>(c =>
        {
            c.Map(p => p.Name).WithColumnName("Full Name");
            c.Map(p => p.Age).WithColumnName("Years");
        });

        var row = importer.ReadSheet().ReadRows<GetOnlyClass>().Single();
        Assert.Equal("Alice", row.Name);
        Assert.Equal(30, row.Age);
    }

    [Fact]
    public void ReadRows_CustomMappedPositionalRecord_Success()
    {
        using var importer = Helpers.GetCsvImporter("Full Name,Years\nAlice,30\n");
        importer.Configuration.RegisterClassMap<PositionalRecord>(c =>
        {
            c.Map(p => p.Name).WithColumnName("Full Name");
            c.Map(p => p.Age).WithColumnName("Years");
        });

        var row = importer.ReadSheet().ReadRows<PositionalRecord>().Single();
        Assert.Equal(new PositionalRecord("Alice", 30), row);
    }

    [Fact]
    public void Map_ReadOnlyPropertyNotConstructorParameter_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<GetOnlyClass>();
        Assert.Throws<ArgumentException>("member", () => map.Map(p => p.Description));
    }

    public class ValidatingClass
    {
        public ValidatingClass(int age)
        {
            if (age < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(age), "Age cannot be negative.");
            }

            Age = age;
        }

        public int Age { get; }
    }

    [Fact]
    public void ReadRows_ConstructorThrows_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetCsvImporter("Age\n1\n-1\n");
        var sheet = importer.ReadSheet();

        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<ValidatingClass>().ToArray());
        Assert.IsType<ArgumentOutOfRangeException>(ex.InnerException);
        Assert.Equal(2, ex.RowIndex);
        Assert.Contains("Age cannot be negative.", ex.Message);
        Assert.Contains("on row 3", ex.Message);
    }

    public class NoMatchingConstructorClass
    {
        public NoMatchingConstructorClass(int somethingElse)
        {
        }

        public string? Name { get; set; }
    }

    [Fact]
    public void ReadRows_NoMatchingConstructor_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetCsvImporter(PeopleCsv);
        importer.Configuration.RegisterClassMap<NoMatchingConstructorClass>(c => c.Map(p => p.Name));
        var sheet = importer.ReadSheet();

        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<NoMatchingConstructorClass>().ToArray());
        Assert.Contains("public parameterless constructor", ex.Message);
    }

    [Fact]
    public void ReadRows_AutoMappedNoMatchingConstructor_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetCsvImporter(PeopleCsv);
        var sheet = importer.ReadSheet();
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<NoMatchingConstructorClass>().ToArray());
    }

    public class Outer
    {
        public PositionalRecord Person { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_CustomMappedNestedRecord_Success()
    {
        using var importer = Helpers.GetCsvImporter(PeopleCsv);
        importer.Configuration.RegisterClassMap<Outer>(c =>
        {
            c.Map(p => p.Person.Name);
            c.Map(p => p.Person.Age);
        });

        var row = importer.ReadSheet().ReadRows<Outer>().First();
        Assert.Equal(new PositionalRecord("Alice", 30), row.Person);
    }

    [Fact]
    public void Properties_ChangedAfterReading_ThrowsInvalidOperationException()
    {
        using var importer = Helpers.GetCsvImporter(PeopleCsv);
        var map = new ExcelClassMap<MultipleConstructorsClass>();
        map.Map(p => p.Name);
        importer.Configuration.RegisterClassMap(map);
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var first = sheet.ReadRow<MultipleConstructorsClass>();
        Assert.Equal(-1, first.Age);

        // The class map is frozen, so the constructor chosen for it stays the same.
        Assert.Throws<InvalidOperationException>(() => map.Map(p => p.Age));
        var second = sheet.ReadRow<MultipleConstructorsClass>();
        Assert.Equal("Bob", second.Name);
        Assert.Equal(-1, second.Age);
    }

    public record AttributedRecord(
        [property: ExcelColumnName("Full Name")] string Name,
        [property: ExcelOptional] decimal Salary = 0);

    [Fact]
    public void ReadRows_RecordWithPropertyAttributes_Success()
    {
        using var importer = Helpers.GetCsvImporter("Full Name,Other\nAlice,x\n");
        var row = importer.ReadSheet().ReadRows<AttributedRecord>().Single();
        Assert.Equal(new AttributedRecord("Alice", 0), row);
    }

    [Fact]
    public void ReadRows_NoUsableConstructor_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetCsvImporter("Value\n1\n");
        importer.Configuration.RegisterClassMap<NoUsableConstructor<int>>(c => c.Map(p => p.Value));
        var sheet = importer.ReadSheet();

        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<NoUsableConstructor<int>>().ToArray());
        Assert.Equal("Cannot create an instance of type MapConstructorTests.NoUsableConstructor<int>. It needs a public parameterless constructor, or a public constructor whose parameters each match a mapped property or field (by name and type) or have a default value.", exception.Message);
    }

    public class NoUsableConstructor<T>(string other)
    {
        public string Other { get; } = other;

        public T? Value { get; set; }
    }
}

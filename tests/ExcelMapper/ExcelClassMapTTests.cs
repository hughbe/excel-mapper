using System.Linq;
﻿
namespace ExcelMapper.Tests;

public class ExcelClassMapTTests
{
    [Theory]
    [InlineData(EmptyValueStrategy.Error - 1)]
    [InlineData(EmptyValueStrategy.DefaultValue + 1)]
    public void Ctor_InvalidEmptyValueStrategy_ThrowsArgumentOutOfRangeException(EmptyValueStrategy emptyValueStrategy)
    {
        Assert.Throws<ArgumentOutOfRangeException>("emptyValueStrategy", () => new TestClassMap(emptyValueStrategy));
    }

    [Fact]
    public void Map_EnumIgnoreCaseNotEnum_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<TestClass>();
    }

    [Fact]
    public void Map_IEnumerable_ThrowsExcelMappingException()
    {
        using var stream = Helpers.GetResource("Primitives.xlsx");
        using var importer = new ExcelImporter(stream);

        var map = new ExcelClassMap<Helpers.TestClass>();
        map.Map(p => p.ConcreteIEnumerable);
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap(map));
    }

    [Fact]
    public void Map_IDictionary_ThrowsExcelMappingException()
    {
        using var stream = Helpers.GetResource("Primitives.xlsx");
        using var importer = new ExcelImporter(stream);

        var map = new ExcelClassMap<Helpers.TestClass>();
        map.Map(p => p.ConcreteIDictionary);
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap(map));
    }

    [Fact]
    public void Map_IDictionaryNoConstructor_ThrowsExcelMappingException()
    {
        using var stream = Helpers.GetResource("Primitives.xlsx");
        using var importer = new ExcelImporter(stream);

        var map = new ExcelClassMap<Helpers.TestClass>();
        map.Map(p => p.IDictionaryNoConstructor);
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap(map));
    }

    [Fact]
    public void MultiMap_UnknownInterface_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<Helpers.TestClass>();
        Assert.Throws<ExcelMappingException>(() => map.Map<string>(p => p.UnknownInterfaceValue));
    }

    [Fact]
    public void MultiMap_ConcreteIEnumerable_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<Helpers.TestClass>();
        Assert.Throws<ExcelMappingException>(() => map.Map<string>(p => p.ConcreteIEnumerable));
    }

    [Fact]
    public void MultiMap_CantMapIEnumerableElementType_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<Helpers.TestClass>();
        Assert.Throws<ExcelMappingException>(() => map.Map(p => p.CantMapElementType));
    }

    [Fact]
    public void MapObject_String_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<Helpers.TestClass>();
        Assert.Throws<ExcelMappingException>(() => map.MapObject(p => p.Value));
    }

    [Fact]
    public void MapObject_Interface_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<Helpers.TestClass>();
        Assert.Throws<ExcelMappingException>(() => map.MapObject(p => p.UnknownInterfaceValue));
    }

    [Fact]
    public void MapObject_InvalidIListMemberType_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<Helpers.TestClass>();
        Assert.Throws<ExcelMappingException>(() => map.MapObject(p => p.InvalidIListMemberType));
    }

    [Fact]
    public void MapObject_InvalidIDictionaryMemberType_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<Helpers.TestClass>();
        Assert.Throws<ExcelMappingException>(() => map.MapObject(p => p.InvalidIDictionaryMemberType));
    }

    [Fact]
    public void Map_InvalidMethodExpression_ThrowsArgumentException()
    {
        var otherType = new OtherType();
        var map = new ExcelClassMap<OtherType>();
        Assert.Throws<ArgumentException>("expression", () => map.Map(p => otherType.Value.ToString()));
    }

    [Fact]
    public void Map_InvalidCastExpression_ThrowsExcelMappingException()
    {
        using var stream = Helpers.GetResource("Primitives.xlsx");
        using var importer = new ExcelImporter(stream);

        var map = new ExcelClassMap<Helpers.TestClass>();
        map.Map(p => (CollectionAttribute)p.ObjectValue);
        Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap(map));
    }

    public class OtherType
    {
        public int Value { get; set; }
    }

    [Fact]
    public void WithClassMap_InvokeClassMapFactory_Success()
    {
        var map = new ExcelClassMap<TestClass>();
        Assert.Same(map, map.WithClassMap(c =>
        {
            Assert.Same(map, c);
        }));
    }

    [Fact]
    public void WithClassMap_NullConfigure_ThrowsArgumentNullException()
    {
        var map = new ExcelClassMap<TestClass>();
        Assert.Throws<ArgumentNullException>("configure", () => map.WithClassMap((Action<ExcelClassMap<TestClass>>)null!));
    }

    [Fact]
    public void WithClassMap_InvokeSameClassMap_Success()
    {
        var map = new ExcelClassMap<TestClass>();
        Assert.Same(map, map.WithClassMap(map));
    }

    [Fact]
    public void MapExtraColumns_Parameter_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<Dictionary<string, string>>();
        Assert.Throws<ArgumentException>("expression", () => map.MapExtraColumns(p => p));
    }

    [Fact]
    public void WithClassMap_SingleValueClassMap_ReadsValue()
    {
        var inner = new ExcelClassMap<int>();
        inner.Map(p => p);
        var outer = new ExcelClassMap<int>();
        Assert.Same(outer, outer.WithClassMap(inner));

        using var importer = Helpers.GetCsvImporter("A\n5\n");
        Assert.Equal([5], importer.ReadSheet().ReadRows(outer).ToArray());
    }

    [Fact]
    public void WithClassMap_NullClassMap_ThrowsArgumentNullException()
    {
        var map = new ExcelClassMap<TestClass>();
        Assert.Throws<ArgumentNullException>("classMap", () => map.WithClassMap((ExcelClassMap<TestClass>)null!));
    }

    private class TestClass
    {
        public int IntValue { get; set; }
        public string StringValue { get; set; } = default!;
        public TestEnum EnumValue { get; set; }
        public TestEnum? NullableEnumValue { get; set; }
        public DateTime DateTimeValue { get; set; }
        public DateTime? NullableDateTimeValue { get; set; }
        public TimeSpan TimeSpanValue { get; set; }
        public TimeSpan? NullableTimeSpanValue { get; set; }
        public ChildClass? ObjectValue { get; set; }
        public long? NullableLongValue { get; set; }
        public Dictionary<string, string> DictionaryValue { get; set; } = default!;
    }

    private enum TestEnum
    {
        Value1,
        Value2
    }

    private class ChildClass
    {
        public int Value { get; set; }
        public Dictionary<string, int> Extra { get; set; } = default!;
    }

    private class TestClassMap : ExcelClassMap<Helpers.TestClass>
    {
        public TestClassMap(EmptyValueStrategy emptyValueStrategy) : base(emptyValueStrategy) { }
    }

    [Fact]
    public void AutoMap_Invoke_MapsEveryMember()
    {
        var map = new ExcelClassMap<AutoMapped>();
        Assert.Same(map, map.AutoMap());

        // Calling it again maps nothing more.
        Assert.Same(map, map.AutoMap());

        using var importer = Helpers.GetCsvImporter("Name,Age,Team\na,1,Finance\n");
        importer.Configuration.RegisterClassMap(map);
        var row = Assert.Single(importer.ReadSheet().ReadRows<AutoMapped>());
        Assert.Equal("a", row.Name);
        Assert.Equal(1, row.Age);
        Assert.Equal(new Dictionary<string, string> { { "Team", "Finance" } }, row.Other);
    }

    [Fact]
    public void AutoMap_MemberAlreadyMapped_KeepsIt()
    {
        // Name keeps its column: auto-mapping it would read the "Name" column, which is missing.
        using var importer = Helpers.GetCsvImporter("Full Name,Age\na,1\n");
        importer.Configuration.RegisterClassMap<AutoMapped>(c =>
        {
            c.Map(r => r.Name).WithColumnName("Full Name");
            c.AutoMap();
        });

        var row = Assert.Single(importer.ReadSheet().ReadRows<AutoMapped>());
        Assert.Equal("a", row.Name);
        Assert.Equal(1, row.Age);
        Assert.Empty(row.Other);
    }

    [Fact]
    public void AutoMap_TypeCantBeAutoMapped_ThrowsExcelMappingException()
    {
        var map = new ExcelClassMap<NotAutoMappable>();
        var exception = Assert.Throws<ExcelMappingException>(() => map.AutoMap());
        Assert.Contains("Cannot auto-map type", exception.Message);
    }

    [Fact]
    public void ReadRows_AutoMapThenMap_ReplacesTheMember()
    {
        // The extra columns exclude the column the replaced member reads, so they belong to this
        // class map rather than one auto-mapping created.
        using var importer = Helpers.GetCsvImporter("Full Name,Age,Team,Desk\na,1,Finance,4B\n");
        importer.Configuration.RegisterClassMap<AutoMapped>(c =>
        {
            c.AutoMap();
            c.Map(r => r.Name).WithColumnName("Full Name");
        });

        var row = Assert.Single(importer.ReadSheet().ReadRows<AutoMapped>());
        Assert.Equal("a", row.Name);
        Assert.Equal(1, row.Age);
        Assert.Equal(new Dictionary<string, string> { { "Team", "Finance" }, { "Desk", "4B" } }, row.Other);
    }

    [Fact]
    public void ReadRows_AutoMapNestedObject_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Name,Street,Zip\na,High St,123\n");
        importer.Configuration.RegisterClassMap<WithAddress>(c =>
        {
            c.Map(r => r.Name);
            c.MapObject(r => r.Address).AutoMap();
        });

        var row = Assert.Single(importer.ReadSheet().ReadRows<WithAddress>());
        Assert.Equal("a", row.Name);
        Assert.Equal("High St", row.Address.Street);
        Assert.Equal(123, row.Address.Zip);
    }

    [Fact]
    public void ReadRows_AutoMapValue_ReadsTheFirstColumn()
    {
        using var importer = Helpers.GetCsvImporter("Value\na\nb\n");
        importer.Configuration.RegisterClassMap<string>(c => c.AutoMap());
        Assert.Equal(["a", "b"], importer.ReadSheet().ReadRows<string>());
    }

    public class AutoMapped
    {
        public string Name { get; set; } = default!;
        public int Age { get; set; }

        [ExcelExtraColumns]
        public Dictionary<string, string> Other { get; set; } = default!;
    }

    public class NotAutoMappable
    {
        public IDisposable Value { get; set; } = default!;
    }

    public class WithAddress
    {
        public string Name { get; set; } = default!;
        public Address Address { get; set; } = default!;
    }

    public class Address
    {
        public string Street { get; set; } = default!;
        public int Zip { get; set; }
    }

    [Fact]
    public void Map_SameMemberTwice_ConfiguresOneMap()
    {
        var classMap = new ExcelClassMap<ReuseRow>();
        var first = classMap.Map(r => r.Name).WithColumnName("Full Name");
        var second = classMap.Map(r => r.Name).MakeOptional();
        Assert.Same(first, second);

        using var importer = Helpers.GetCsvImporter("Full Name\nAlice\n");
        importer.Configuration.RegisterClassMap(classMap);
        var row = importer.ReadSheet().ReadRows<ReuseRow>().Single();
        Assert.Equal("Alice", row.Name);
        Assert.True(first.Optional);
    }

    [Fact]
    public void Map_AfterAutoMap_ReplacesAutoMappedMapThenReusesIt()
    {
        var classMap = new ExcelClassMap<NameRow>();
        classMap.AutoMap();
        var mapped = classMap.Map(r => r.Name).WithColumnName("Full Name");
        Assert.Same(mapped, classMap.Map(r => r.Name));

        // The auto-mapped map, which reads "Name", is replaced.
        using var importer = Helpers.GetCsvImporter("Full Name\nAlice\n");
        Assert.Equal("Alice", importer.ReadSheet().ReadRows(classMap).Single().Name);
    }

    public class NameRow
    {
        public string? Name { get; set; }
    }

    [Fact]
    public void Map_SameMemberAsAnotherKindOfMap_ReplacesMap()
    {
        // Map gives a value map that can't read a HashSet; MapCollection replaces it.
        var classMap = new ExcelClassMap<TagsRow>();
        classMap.Map(r => r.Tags);
        classMap.MapCollection<string>(r => r.Tags);

        using var importer = Helpers.GetCsvImporter("Tags\n\"a,b\"\n");
        Assert.Equal(["a", "b"], importer.ReadSheet().ReadRows(classMap).Single().Tags.Order());
    }

    public class TagsRow
    {
        public HashSet<string> Tags { get; set; } = [];
    }

    [Fact]
    public void Map_NestedMemberAndIndexerTwice_ReuseMaps()
    {
        var classMap = new ExcelClassMap<ReuseRow>();
        Assert.Same(classMap.Map(r => r.Address.Street), classMap.Map(r => r.Address.Street));
        Assert.Same(classMap.Map(r => r.Scores[0]), classMap.Map(r => r.Scores[0]));
        Assert.NotSame(classMap.Map(r => r.Scores[0]), classMap.Map(r => r.Scores[1]));
    }

    [Fact]
    public void MapObject_AfterNestedMember_KeepsNestedMember()
    {
        var classMap = new ExcelClassMap<AddressRow>();
        classMap.Map(r => r.Address.Street).WithColumnName("Road");
        classMap.MapObject(r => r.Address);

        using var importer = Helpers.GetCsvImporter("Road\nHigh Street\n");
        Assert.Equal("High Street", importer.ReadSheet().ReadRows(classMap).Single().Address.Street);
    }

    public class AddressRow
    {
        public ReuseAddress Address { get; set; } = new();
    }

    [Fact]
    public void Map_AfterReadingRows_ThrowsInvalidOperationException()
    {
        var classMap = new ExcelClassMap<ReuseRow>();
        classMap.Map(r => r.Name);
        using var importer = Helpers.GetCsvImporter("Name\nAlice\n");
        _ = importer.ReadSheet().ReadRows(classMap).ToArray();

        // Even a member the class map already maps: its map can't change either.
        Assert.Throws<InvalidOperationException>(() => classMap.Map(r => r.Name));
    }

    public class ReuseRow
    {
        public string? Name { get; set; }

        public HashSet<string> Tags { get; set; } = [];

        public ReuseAddress Address { get; set; } = new();

        public int[] Scores { get; set; } = [];
    }

    public class ReuseAddress
    {
        public string? Street { get; set; }
    }
}

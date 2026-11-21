
namespace ExcelMapper.Tests;

public class ExcelClassMapTTests
{
    [Theory]
    [InlineData(FallbackStrategy.ThrowIfPrimitive - 1)]
    [InlineData(FallbackStrategy.SetToDefaultValue + 1)]
    public void Ctor_InvalidEmptyValueStrategy_ThrowsArgumentOutOfRangeException(FallbackStrategy emptyValueStrategy)
    {
        Assert.Throws<ArgumentOutOfRangeException>("emptyValueStrategy", () => new TestClassMap(emptyValueStrategy));
    }

    [Fact]
    public void Map_EnumIgnoreCaseNotEnum_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<TestClass>();
        Assert.Throws<ArgumentException>("TProperty", () => map.Map(p => p.DateTimeValue, ignoreCase: true));
        Assert.Throws<ArgumentException>("TProperty", () => map.Map(p => p.NullableDateTimeValue, ignoreCase: true));
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
    public void WithClassMap_NullClassMapFactory_ThrowsArgumentNullException()
    {
        var map = new ExcelClassMap<TestClass>();
        Assert.Throws<ArgumentNullException>("classMapFactory", () => map.WithClassMap((Action<ExcelClassMap<TestClass>>)null!));
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
        public TestClassMap(FallbackStrategy emptyValueStrategy) : base(emptyValueStrategy) { }
    }
}

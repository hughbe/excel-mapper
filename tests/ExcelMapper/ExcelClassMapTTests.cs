using ExcelMapper.Mappers;
using ExcelMapper.Readers;

namespace ExcelMapper.Tests;

public class ExcelClassMapTTests
{
    [Fact]
    public void Ctor_Default()
    {
        var map = new ExcelClassMap<string>();
        Assert.Equal(typeof(string), map.Type);
        Assert.Equal(FallbackStrategy.ThrowIfPrimitive, map.EmptyValueStrategy);
        Assert.Empty(map.Properties);
        Assert.Same(map.Properties, map.Properties);
    }

    [Theory]
    [InlineData(FallbackStrategy.ThrowIfPrimitive)]
    [InlineData(FallbackStrategy.SetToDefaultValue)]
    public void Ctor_EmptyValueStrategy(FallbackStrategy emptyValueStrategy)
    {
        var map = new ExcelClassMap<string>(emptyValueStrategy);
        Assert.Equal(emptyValueStrategy, map.EmptyValueStrategy);
        Assert.Equal(typeof(string), map.Type);
        Assert.Empty(map.Properties);
        Assert.Same(map.Properties, map.Properties);
    }

    [Theory]
    [InlineData(FallbackStrategy.ThrowIfPrimitive - 1)]
    [InlineData(FallbackStrategy.SetToDefaultValue + 1)]
    public void Ctor_InvalidEmptyValueStrategy_ThrowsArgumentOutOfRangeException(FallbackStrategy emptyValueStrategy)
    {
        Assert.Throws<ArgumentOutOfRangeException>("emptyValueStrategy", () => new TestClassMap(emptyValueStrategy));
    }

    [Fact]
    public void Map_FuncExpressionT_Success()
    {
        var map = new ExcelClassMap<TestClass>();
        var result = map.Map(p => p.StringValue);
        Assert.Single(map.Properties);
        Assert.Equal(typeof(TestClass).GetProperty(nameof(TestClass.StringValue)), map.Properties[0].Member);
        Assert.Same(result, map.Properties[0].Map);

        // Map again.
        var result2 = map.Map(p => p.StringValue);
        Assert.NotSame(result, result2);
        Assert.Single(map.Properties);
        Assert.Equal(typeof(TestClass).GetProperty(nameof(TestClass.StringValue)), map.Properties[0].Member);
        Assert.Same(result2, map.Properties[0].Map);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Map_FuncExpressionTEnumBool_Success(bool ignoreCase)
    {
        var map = new ExcelClassMap<TestClass>();
        var result = map.Map(p => p.EnumValue, ignoreCase);
        Assert.Single(map.Properties);
        Assert.Equal(typeof(TestClass).GetProperty(nameof(TestClass.EnumValue)), map.Properties[0].Member);
        Assert.Same(result, map.Properties[0].Map);
        var mapper = Assert.IsAssignableFrom<EnumMapper>(Assert.Single(result.Mappers));
        Assert.Equal(typeof(TestEnum), mapper.EnumType);
        Assert.Equal(ignoreCase, mapper.IgnoreCase);

        // Map again.
        var result2 = map.Map(p => p.EnumValue, ignoreCase);
        Assert.NotSame(result, result2);
        Assert.Single(map.Properties);
        Assert.Equal(typeof(TestClass).GetProperty(nameof(TestClass.EnumValue)), map.Properties[0].Member);
        Assert.Same(result2, map.Properties[0].Map);
        mapper = Assert.IsAssignableFrom<EnumMapper>(Assert.Single(result2.Mappers));
        Assert.Equal(typeof(TestEnum), mapper.EnumType);
        Assert.Equal(ignoreCase, mapper.IgnoreCase);

        // Map opposite.
        var result3 = map.Map(p => p.EnumValue, !ignoreCase);
        Assert.NotSame(result2, result3);
        Assert.Single(map.Properties);
        Assert.Equal(typeof(TestClass).GetProperty(nameof(TestClass.EnumValue)), map.Properties[0].Member);
        Assert.Same(result3, map.Properties[0].Map);
        mapper = Assert.IsAssignableFrom<EnumMapper>(Assert.Single(result3.Mappers));
        Assert.Equal(typeof(TestEnum), mapper.EnumType);
        Assert.Equal(!ignoreCase, mapper.IgnoreCase);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Map_FuncExpressionNullableTEnumBool_Success(bool ignoreCase)
    {
        var map = new ExcelClassMap<TestClass>();
        var result = map.Map(p => p.NullableEnumValue, ignoreCase);
        Assert.Single(map.Properties);
        Assert.Equal(typeof(TestClass).GetProperty(nameof(TestClass.NullableEnumValue)), map.Properties[0].Member);
        Assert.Same(result, map.Properties[0].Map);
        var mapper = Assert.IsAssignableFrom<EnumMapper>(Assert.Single(result.Mappers));
        Assert.Equal(typeof(TestEnum), mapper.EnumType);
        Assert.Equal(ignoreCase, mapper.IgnoreCase);

        // Map again.
        var result2 = map.Map(p => p.NullableEnumValue, ignoreCase);
        Assert.NotSame(result, result2);
        Assert.Single(map.Properties);
        Assert.Equal(typeof(TestClass).GetProperty(nameof(TestClass.NullableEnumValue)), map.Properties[0].Member);
        Assert.Same(result2, map.Properties[0].Map);
        mapper = Assert.IsAssignableFrom<EnumMapper>(Assert.Single(result2.Mappers));
        Assert.Equal(typeof(TestEnum), mapper.EnumType);
        Assert.Equal(ignoreCase, mapper.IgnoreCase);

        // Map opposite.
        var result3 = map.Map(p => p.NullableEnumValue, !ignoreCase);
        Assert.NotSame(result2, result3);
        Assert.Single(map.Properties);
        Assert.Equal(typeof(TestClass).GetProperty(nameof(TestClass.NullableEnumValue)), map.Properties[0].Member);
        Assert.Same(result3, map.Properties[0].Map);
        mapper = Assert.IsAssignableFrom<EnumMapper>(Assert.Single(result3.Mappers));
        Assert.Equal(typeof(TestEnum), mapper.EnumType);
        Assert.Equal(!ignoreCase, mapper.IgnoreCase);
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
    public void MapObject_ClassMapFactory_ReturnsExpected()
    {
        var map = new TestClassMap(FallbackStrategy.ThrowIfPrimitive);
        var classMap = map.MapObject(t => t.ObjectValue);
        Assert.Empty(classMap.Properties);
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
    public void WithClassMap_InvokeClassMap_Success()
    {
        var map = new ExcelClassMap<TestClass>();
        map.Map(p => p.EnumValue);
        var otherMap = new ExcelClassMap<TestClass>();
        otherMap.Map(p => p.IntValue);
        otherMap.Map(p => p.StringValue);
        Assert.Same(map, map.WithClassMap(otherMap));
        Assert.Equal(2, map.Properties.Count);
        Assert.Equal(nameof(TestClass.IntValue), map.Properties[0].Member.Name);
        Assert.Equal(nameof(TestClass.StringValue), map.Properties[1].Member.Name);
    }

    [Fact]
    public void WithClassMap_InvokeClassMapEmpty_Success()
    {
        var map = new ExcelClassMap<TestClass>();
        map.Map(p => p.EnumValue);
        var otherMap = new ExcelClassMap<TestClass>();
        Assert.Same(map, map.WithClassMap(otherMap));
        Assert.Empty(map.Properties);
    }

    [Fact]
    public void WithClassMap_InvokeSameClassMap_Success()
    {
        var map = new ExcelClassMap<TestClass>();
        Assert.Same(map, map.WithClassMap(map));
    }

    [Fact]
    public void MapRowNumber_FuncExpressionT_Success()
    {
        var map = new ExcelClassMap<TestClass>();
        var result = map.MapRowNumber(p => p.IntValue);
        Assert.Equal(typeof(int), result.Type);
        Assert.Single(map.Properties);
        Assert.Equal(typeof(TestClass).GetProperty(nameof(TestClass.IntValue)), map.Properties[0].Member);
        Assert.Same(result, map.Properties[0].Map);

        // Map again.
        var result2 = map.MapRowNumber(p => p.IntValue);
        Assert.NotSame(result, result2);
        Assert.Single(map.Properties);
        Assert.Same(result2, map.Properties[0].Map);
    }

    [Fact]
    public void MapRowNumber_NullableMember_Success()
    {
        var map = new ExcelClassMap<TestClass>();
        var result = map.MapRowNumber(p => p.NullableLongValue);
        Assert.Equal(typeof(long?), result.Type);
        Assert.Same(result, Assert.Single(map.Properties).Map);
    }

    [Fact]
    public void MapRowNumber_NestedMember_Success()
    {
        var map = new ExcelClassMap<TestClass>();
        var result = map.MapRowNumber(p => p.ObjectValue!.Value);
        var nestedMap = Assert.IsType<ExcelClassMap<ChildClass>>(Assert.Single(map.Properties).Map);
        Assert.Same(result, Assert.Single(nestedMap.Properties).Map);
    }

    [Fact]
    public void MapRowNumber_InvalidMemberType_ThrowsArgumentException()
    {
        var map = new ExcelClassMap<TestClass>();
        Assert.Throws<ArgumentException>("expression", () => map.MapRowNumber(p => p.StringValue));
        Assert.Throws<ArgumentException>("expression", () => map.MapRowNumber(p => p.EnumValue));
        Assert.Throws<ArgumentException>("expression", () => map.MapRowNumber(p => p.DateTimeValue));
        Assert.Empty(map.Properties);
    }

    [Fact]
    public void MapExtraColumns_FuncExpressionT_Success()
    {
        var map = new ExcelClassMap<TestClass>();
        var result = map.MapExtraColumns(p => p.DictionaryValue);
        Assert.Single(map.Properties);
        Assert.Equal(typeof(TestClass).GetProperty(nameof(TestClass.DictionaryValue)), map.Properties[0].Member);
        Assert.Same(result, map.Properties[0].Map);

        // Map again.
        var result2 = map.MapExtraColumns(p => p.DictionaryValue);
        Assert.NotSame(result, result2);
        Assert.Single(map.Properties);
        Assert.Same(result2, map.Properties[0].Map);
    }

    [Fact]
    public void MapExtraColumns_NestedMember_Success()
    {
        var map = new ExcelClassMap<TestClass>();
        var result = map.MapExtraColumns(p => p.ObjectValue!.Extra);
        var nestedMap = Assert.IsType<ExcelClassMap<ChildClass>>(Assert.Single(map.Properties).Map);
        Assert.Same(result, Assert.Single(nestedMap.Properties).Map);
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

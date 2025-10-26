using ExcelMapper.Readers;

namespace ExcelMapper.Tests;

public class ExcelClassMapTests
{
    [Fact]
    public void Ctor_Type()
    {
        var map = new ExcelClassMap(typeof(string));
        Assert.Equal(typeof(string), map.Type);
        Assert.Empty(map.Properties);
        Assert.Same(map.Properties, map.Properties);
        Assert.Equal(FallbackStrategy.ThrowIfPrimitive, map.EmptyValueStrategy);
    }

    [Theory]
    [InlineData(FallbackStrategy.ThrowIfPrimitive)]
    [InlineData(FallbackStrategy.SetToDefaultValue)]
    public void Ctor_Type_FallbackStrategy(FallbackStrategy emptyValueStrategy)
    {
        var map = new ExcelClassMap(typeof(string), emptyValueStrategy);
        Assert.Equal(typeof(string), map.Type);
        Assert.Empty(map.Properties);
        Assert.Same(map.Properties, map.Properties);
        Assert.Equal(emptyValueStrategy, map.EmptyValueStrategy);
    }

    [Fact]
    public void Ctor_NullType_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("type", () => new ExcelClassMap(null!));
        Assert.Throws<ArgumentNullException>("type", () => new ExcelClassMap(null!, FallbackStrategy.ThrowIfPrimitive));
    }

    [Theory]
    [InlineData(FallbackStrategy.ThrowIfPrimitive - 1)]
    [InlineData(FallbackStrategy.SetToDefaultValue + 1)]
    public void Ctor_InvalidEmptyValueStrategy_ThrowsArgumentOutOfRangeException(FallbackStrategy emptyValueStrategy)
    {
        Assert.Throws<ArgumentOutOfRangeException>("emptyValueStrategy", () => new ExcelClassMap(typeof(string), emptyValueStrategy));
    }

    [Fact]
    public void Properties_Get_ReturnsExpected()
    {
        var map = new ExcelClassMap(typeof(string));
        var properties = map.Properties;
        Assert.Empty(properties);
        Assert.Same(properties, map.Properties);
    }

    [Fact]
    public void Properties_AddValidItem_Success()
    {
        var propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Property))!;
        var map1 = new OneToOneMap<int>(new ColumnNameReaderFactory("Property"));
        var propertyMap1 = new ExcelPropertyMap(propertyInfo, map1);
        var map2 = new OneToOneMap<int>(new ColumnNameReaderFactory("Property"));
        var propertyMap2 = new ExcelPropertyMap(propertyInfo, map2);
        var properties = new ExcelClassMap<TestClass>().Properties;

        properties.Add(propertyMap1);
        Assert.Same(propertyMap1, Assert.Single(properties));
        Assert.Same(propertyMap1, properties[0]);

        properties.Add(propertyMap2);
        Assert.Equal(2, properties.Count);
        Assert.Same(propertyMap2, properties[1]);

        properties.Add(propertyMap1);
        Assert.Equal(3, properties.Count);
        Assert.Same(propertyMap1, properties[2]);
    }

    [Fact]
    public void Properties_AddNullItem_ThrowsArgumentNullException()
    {
        var properties = new ExcelClassMap<TestClass>().Properties;
        Assert.Throws<ArgumentNullException>("item", () => properties.Add(null!));
    }

    [Fact]
    public void Properties_InsertValidItem_Success()
    {
        var propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Property))!;
        var map1 = new OneToOneMap<int>(new ColumnNameReaderFactory("Property"));
        var propertyMap1 = new ExcelPropertyMap(propertyInfo, map1);
        var map2 = new OneToOneMap<int>(new ColumnNameReaderFactory("Property"));
        var propertyMap2 = new ExcelPropertyMap(propertyInfo, map2);
        var properties = new ExcelClassMap<TestClass>().Properties;

        properties.Insert(0, propertyMap1);
        Assert.Same(propertyMap1, Assert.Single(properties));
        Assert.Same(propertyMap1, properties[0]);

        properties.Insert(0, propertyMap2);
        Assert.Equal(2, properties.Count);
        Assert.Same(propertyMap2, properties[0]);

        properties.Insert(1, propertyMap1);
        Assert.Equal(3, properties.Count);
        Assert.Same(propertyMap1, properties[1]);
    }

    [Fact]
    public void Properties_InsertNullItem_ThrowsArgumentNullException()
    {
        var properties = new ExcelClassMap<TestClass>().Properties;
        Assert.Throws<ArgumentNullException>("item", () => properties.Insert(0, null!));
    }

    [Fact]
    public void Properties_ItemSetValidItem_GetReturnsExpected()
    {
        var propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Property))!;
        var map1 = new OneToOneMap<int>(new ColumnNameReaderFactory("Property"));
        var propertyMap1 = new ExcelPropertyMap(propertyInfo, map1);
        var map2 = new OneToOneMap<int>(new ColumnNameReaderFactory("Property"));
        var propertyMap2 = new ExcelPropertyMap(propertyInfo, map2);
        var properties = new ExcelClassMap<TestClass>().Properties;
        properties.Add(propertyMap1);

        properties[0] = propertyMap2;
        Assert.Same(propertyMap2, properties[0]);
    }

    [Fact]
    public void Properties_ItemSetNull_ThrowsArgumentNullException()
    {
        var propertyInfo = typeof(TestClass).GetProperty(nameof(TestClass.Property))!;
        var map = new OneToOneMap<int>(new ColumnNameReaderFactory("Property"));
        var propertyMap = new ExcelPropertyMap(propertyInfo, map);
        var properties = new ExcelClassMap<TestClass>().Properties;
        properties.Add(propertyMap);

        Assert.Throws<ArgumentNullException>("item", () => properties[0] = null!);
    }


    private class TestClass
    {
        public int Property { get; set; }
    }
}

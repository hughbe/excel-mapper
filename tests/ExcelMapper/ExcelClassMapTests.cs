using System;
using Xunit;

namespace ExcelMapper.Tests;

public class ExcelClassMapTests : ExcelClassMap<Helpers.TestClass>
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
    public void Ctor_InvalidEmptyValueStrategy_ThrowsArgumentException(FallbackStrategy emptyValueStrategy)
    {
        Assert.Throws<ArgumentException>("emptyValueStrategy", () => new ExcelClassMap(typeof(string), emptyValueStrategy));
    }
}

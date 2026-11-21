
namespace ExcelMapper.Tests;

public class ExcelClassMapTests
{
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


    private class TestClass
    {
        public int Property { get; set; }
    }
}

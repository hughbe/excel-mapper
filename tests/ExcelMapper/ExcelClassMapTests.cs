
namespace ExcelMapper.Tests;

public class ExcelClassMapTests
{
    [Fact]
    public void Ctor_NullType_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("type", () => new ExcelClassMap(null!));
        Assert.Throws<ArgumentNullException>("type", () => new ExcelClassMap(null!, EmptyValueStrategy.Error));
    }

    [Theory]
    [InlineData(EmptyValueStrategy.Error - 1)]
    [InlineData(EmptyValueStrategy.DefaultValue + 1)]
    public void Ctor_InvalidEmptyValueStrategy_ThrowsArgumentOutOfRangeException(EmptyValueStrategy emptyValueStrategy)
    {
        Assert.Throws<ArgumentOutOfRangeException>("emptyValueStrategy", () => new ExcelClassMap(typeof(string), emptyValueStrategy));
    }


    private class TestClass
    {
        public int Property { get; set; }
    }
}

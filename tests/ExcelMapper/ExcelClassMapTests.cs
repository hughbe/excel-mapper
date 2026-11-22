
namespace ExcelMapper.Tests;

public class ExcelClassMapTests
{
    [Theory]
    [InlineData(EmptyValueStrategy.Error - 1)]
    [InlineData(EmptyValueStrategy.DefaultValue + 1)]
    public void Ctor_InvalidEmptyValueStrategy_ThrowsArgumentOutOfRangeException(EmptyValueStrategy emptyValueStrategy)
    {
        Assert.Throws<ArgumentOutOfRangeException>("emptyValueStrategy", () => new ExcelClassMap<string>(emptyValueStrategy));
    }


    private class TestClass
    {
        public int Property { get; set; }
    }
}

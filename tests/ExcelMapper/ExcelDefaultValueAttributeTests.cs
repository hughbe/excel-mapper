namespace ExcelMapper.Tests;

public class ExcelDefaultValueAttributeTests
{
    [Theory]
    [InlineData(null)]
    [InlineData("value")]
    [InlineData(1)]
    public void Ctor_Default(object? value)
    {
        var attribute = new ExcelDefaultValueAttribute(value);
        Assert.Equal(value, attribute.Value);
    }
}

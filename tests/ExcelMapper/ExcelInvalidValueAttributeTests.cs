namespace ExcelMapper.Tests;

public class ExcelInvalidValueAttributeTests
{
    [Theory]
    [InlineData(null)]
    [InlineData("value")]
    [InlineData(1)]
    public void Ctor_Object(object? value)
    {
        var attribute = new ExcelInvalidValueAttribute(value);
        Assert.Equal(value, attribute.Value);
    }
}

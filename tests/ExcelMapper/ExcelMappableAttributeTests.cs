namespace ExcelMapper.Tests;

public class ExcelMappableAttributeTests
{
    [Theory]
    [InlineData(typeof(int))]
    [InlineData(typeof(ExcelMappableAttributeTests))]
    public void Ctor_Type(Type type)
    {
        var attribute = new ExcelMappableAttribute(type);
        Assert.Same(type, attribute.Type);
    }

    [Fact]
    public void Ctor_NullType_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("type", () => new ExcelMappableAttribute(null!));
    }
}

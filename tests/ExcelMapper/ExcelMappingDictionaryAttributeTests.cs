namespace ExcelMapper.Tests;

public class ExcelMappingDictionaryAttributeTests
{
    [Theory]
    [InlineData("value", null)]
    [InlineData("value", "value")]
    [InlineData("value", "mappedValue")]
    [InlineData("value", 2)]
    public void Ctor_Object_Object(string value, object? mappedValue)
    {
        var attribute = new ExcelMappingDictionaryAttribute(value, mappedValue);
        Assert.Equal(value, attribute.Value);
        Assert.Equal(mappedValue, attribute.MappedValue);
    }

    [Fact]
    public void Ctor_NullValue_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("value", () => new ExcelMappingDictionaryAttribute(null!, "mappedValue"));
    }

    [Fact]
    public void Ctor_EmptyValue_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>("value", () => new ExcelMappingDictionaryAttribute(string.Empty, "mappedValue"));
    }
}

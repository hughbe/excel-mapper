
namespace ExcelMapper.Tests;

public class ExcelMappingDictionaryBehaviorAttributeTests
{
    [Theory]
    [InlineData(MappingDictionaryBehavior.Optional)]
    [InlineData(MappingDictionaryBehavior.Required)]
    public void Ctor_MappingDictionary(MappingDictionaryBehavior behavior)
    {
        var attribute = new ExcelMappingDictionaryBehaviorAttribute(behavior);
        Assert.Equal(behavior, attribute.Behavior);
    }

    [Theory]
    [InlineData(MappingDictionaryBehavior.Optional - 1)]
    [InlineData(MappingDictionaryBehavior.Required + 1)]
    public void Ctor_InvalidMappingDictionaryMapperBehavior_ThrowsArgumentOutOfRangeException(MappingDictionaryBehavior behavior)
    {
        Assert.Throws<ArgumentOutOfRangeException>("behavior", () => new ExcelMappingDictionaryBehaviorAttribute(behavior));
    }
}

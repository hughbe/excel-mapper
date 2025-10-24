using System.Runtime.InteropServices;
using ExcelMapper.Mappers;

namespace ExcelMapper.Tests;

public class ExcelMappingDictionaryBehaviorAttributeTests
{
    [Theory]
    [InlineData(MappingDictionaryMapperBehavior.Optional)]
    [InlineData(MappingDictionaryMapperBehavior.Required)]
    public void Ctor_MappingDictionary(MappingDictionaryMapperBehavior behavior)
    {
        var attribute = new ExcelMappingDictionaryBehaviorAttribute(behavior);
        Assert.Equal(behavior, attribute.Behavior);
    }

    [Theory]
    [InlineData(MappingDictionaryMapperBehavior.Optional - 1)]
    [InlineData(MappingDictionaryMapperBehavior.Required + 1)]
    public void Ctor_InvalidMappingDictionaryMapperBehavior_ThrowsArgumentOutOfRangeException(MappingDictionaryMapperBehavior behavior)
    {
        Assert.Throws<ArgumentOutOfRangeException>("behavior", () => new ExcelMappingDictionaryBehaviorAttribute(behavior));
    }
}

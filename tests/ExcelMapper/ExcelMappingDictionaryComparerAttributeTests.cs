namespace ExcelMapper.Tests;

public class ExcelMappingDictionaryComparerAttributeTests
{
    [Theory]
    [InlineData(StringComparison.CurrentCulture)]
    [InlineData(StringComparison.CurrentCultureIgnoreCase)]
    [InlineData(StringComparison.InvariantCulture)]
    [InlineData(StringComparison.InvariantCultureIgnoreCase)]
    [InlineData(StringComparison.Ordinal)]
    [InlineData(StringComparison.OrdinalIgnoreCase)]
    public void Ctor_StringComparison(StringComparison comparison)
    {
        var attribute = new ExcelMappingDictionaryComparerAttribute(comparison);
        Assert.Equal(comparison, attribute.Comparison);
    }

    [Theory]
    [InlineData(StringComparison.CurrentCulture - 1)]
    [InlineData(StringComparison.OrdinalIgnoreCase + 1)]
    public void Ctor_InvalidStringComparison_ThrowsArgumentOutOfRangeException(StringComparison comparison)
    {
        Assert.Throws<ArgumentOutOfRangeException>("comparison", () => new ExcelMappingDictionaryComparerAttribute(comparison));
    }
}

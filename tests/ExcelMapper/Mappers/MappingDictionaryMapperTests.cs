using ExcelMapper.Abstractions;

namespace ExcelMapper.Mappers.Tests;

public class MappingDictionaryMapperTests
{
    public static IEnumerable<object?[]> Ctor_TestData()
    {
        yield return new object?[] { new Dictionary<string, object> { { "key", "value" } }, MappingDictionaryMapperBehavior.Optional };
        yield return new object?[] { new Dictionary<string, object>(), MappingDictionaryMapperBehavior.Optional };
    }

    [Theory]
    [MemberData(nameof(Ctor_TestData))]
    public void Ctor_Dictionary_NullIEqualityComparer_DictionaryBehavior(IDictionary<string, object> mappingDictionary, MappingDictionaryMapperBehavior behavior)
    {
        StringComparer comparer = StringComparer.CurrentCultureIgnoreCase;
        var item = new MappingDictionaryMapper<object>(mappingDictionary, comparer, behavior);

        var itemMapping = Assert.IsType<Dictionary<string, object>>(item.MappingDictionary);
        Assert.Equal(mappingDictionary, itemMapping);
        Assert.Same(comparer, itemMapping.Comparer);
        Assert.Equal(behavior, item.Behavior);
    }

    [Theory]
    [MemberData(nameof(Ctor_TestData))]
    public void Ctor_Dictionary_NonNullIEqualityComparer_DictionaryBehavior(IDictionary<string, object> mappingDictionary, MappingDictionaryMapperBehavior behavior)
    {
        var item = new MappingDictionaryMapper<object>(mappingDictionary, null, behavior);

        var itemMapping = Assert.IsType<Dictionary<string, object>>(item.MappingDictionary);
        Assert.Equal(mappingDictionary, itemMapping);
        Assert.NotNull(itemMapping.Comparer);
        Assert.Equal(behavior, item.Behavior);
    }

    [Fact]
    public void Ctor_NullDictionary_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("mappingDictionary", () => new MappingDictionaryMapper<int>(null!, StringComparer.CurrentCulture, MappingDictionaryMapperBehavior.Required));
    }

    [Theory]
    [InlineData(MappingDictionaryMapperBehavior.Optional - 1)]
    [InlineData(MappingDictionaryMapperBehavior.Required + 1)]
    public void Ctor_InvalidBehavior_ThrowsArgumentOutOfRangeException(MappingDictionaryMapperBehavior behavior)
    {
        Assert.Throws<ArgumentOutOfRangeException>("behavior", () => new MappingDictionaryMapper<object>(new Dictionary<string, object>(), null, behavior));
    }

    [Theory]
    [InlineData(null, false, CellMapperResult.HandleAction.IgnoreResultAndContinueMapping, null)]
    [InlineData("key", true, CellMapperResult.HandleAction.UseResultAndStopMapping, "value")]
    [InlineData("key2", true, CellMapperResult.HandleAction.UseResultAndStopMapping, 10)]
    [InlineData("no_such_key", false, CellMapperResult.HandleAction.IgnoreResultAndContinueMapping, null)]
    public void Map_ValidStringValueOptional_ReturnsSuccess(string? stringValue, bool expectedSucceeded, CellMapperResult.HandleAction expectedAction, object? expectedValue)
    {
        var mapping = new Dictionary<string, object> { { "key", "value" }, { "KEY2", 10 } };
        var comparer = StringComparer.OrdinalIgnoreCase;
        var item = new MappingDictionaryMapper<object>(mapping, comparer, MappingDictionaryMapperBehavior.Optional);

        var result = item.Map(new ReadCellResult(0, stringValue, preserveFormatting: false));
        Assert.Equal(expectedSucceeded, result.Succeeded);
        Assert.Equal(expectedAction, result.Action);
        Assert.Equal(expectedValue, result.Value);
        Assert.Null(result.Exception);
    }

    [Theory]
    [InlineData(null, false, CellMapperResult.HandleAction.IgnoreResultAndContinueMapping, null)]
    [InlineData("key", true, CellMapperResult.HandleAction.UseResultAndStopMapping, "value")]
    [InlineData("key2", true, CellMapperResult.HandleAction.UseResultAndStopMapping, 10)]
    [InlineData("no_such_key", false, CellMapperResult.HandleAction.IgnoreResultAndContinueMapping, null)]
    public void Map_ValidStringValueRequired_ReturnsSuccess(string? stringValue, bool expectedSucceeded, CellMapperResult.HandleAction expectedAction, object? expectedValue)
    {
        var mapping = new Dictionary<string, object> { { "key", "value" }, { "KEY2", 10 } };
        var comparer = StringComparer.OrdinalIgnoreCase;
        var item = new MappingDictionaryMapper<object>(mapping, comparer, MappingDictionaryMapperBehavior.Optional);

        var result = item.Map(new ReadCellResult(0, stringValue, preserveFormatting: false));
        Assert.Equal(expectedSucceeded, result.Succeeded);
        Assert.Equal(expectedAction, result.Action);
        Assert.Equal(expectedValue, result.Value);
        Assert.Null(result.Exception);
    }
}

using System;
using System.Collections.Generic;
using ExcelMapper.Abstractions;
using Xunit;

namespace ExcelMapper.Mappers.Tests;

public class DictionaryMapperTests
{
    public static IEnumerable<object?[]> Ctor_TestData()
    {
        yield return new object?[] { new Dictionary<string, object> { { "key", "value" } }, DictionaryMapperBehavior.Optional };
        yield return new object?[] { new Dictionary<string, object>(), DictionaryMapperBehavior.Optional };
    }

    [Theory]
    [MemberData(nameof(Ctor_TestData))]
    public void Ctor_Dictionary_NullIEqualityComparer_DictionaryBehavior(IDictionary<string, object> mappingDictionary, DictionaryMapperBehavior behavior)
    {
        StringComparer comparer = StringComparer.CurrentCultureIgnoreCase;
        var item = new DictionaryMapper<object>(mappingDictionary, comparer, behavior);

        var itemMapping = Assert.IsType<Dictionary<string, object>>(item.MappingDictionary);
        Assert.Equal(mappingDictionary, itemMapping);
        Assert.Same(comparer, itemMapping.Comparer);
        Assert.Equal(behavior, item.Behavior);
    }

    [Theory]
    [MemberData(nameof(Ctor_TestData))]
    public void Ctor_Dictionary_NonNullIEqualityComparer_DictionaryBehavior(IDictionary<string, object> mappingDictionary, DictionaryMapperBehavior behavior)
    {
        var item = new DictionaryMapper<object>(mappingDictionary, null, behavior);

        var itemMapping = Assert.IsType<Dictionary<string, object>>(item.MappingDictionary);
        Assert.Equal(mappingDictionary, itemMapping);
        Assert.NotNull(itemMapping.Comparer);
        Assert.Equal(behavior, item.Behavior);
    }

    [Fact]
    public void Ctor_NullDictionary_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("mappingDictionary", () => new DictionaryMapper<int>(null!, StringComparer.CurrentCulture, DictionaryMapperBehavior.Required));
    }

    [Theory]
    [InlineData(DictionaryMapperBehavior.Optional - 1)]
    [InlineData(DictionaryMapperBehavior.Required + 1)]
    public void Ctor_InvalidBehavior_ThrowsArgumentException(DictionaryMapperBehavior behavior)
    {
        Assert.Throws<ArgumentException>("behavior", () => new DictionaryMapper<object>(new Dictionary<string, object>(), null, behavior));
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
        var item = new DictionaryMapper<object>(mapping, comparer, DictionaryMapperBehavior.Optional);

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
        var item = new DictionaryMapper<object>(mapping, comparer, DictionaryMapperBehavior.Optional);

        var result = item.Map(new ReadCellResult(0, stringValue, preserveFormatting: false));
        Assert.Equal(expectedSucceeded, result.Succeeded);
        Assert.Equal(expectedAction, result.Action);
        Assert.Equal(expectedValue, result.Value);
        Assert.Null(result.Exception);
    }
}

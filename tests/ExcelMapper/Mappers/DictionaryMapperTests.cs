using System;
using System.Collections.Generic;
using ExcelMapper.Abstractions;
using Xunit;

namespace ExcelMapper.Mappers.Tests;

public class DictionaryMapperTests
{
    [Fact]
    public void Ctor_Dictionary_ReturnsExpected()
    {
        var mapping = new Dictionary<string, object> { { "key", "value" } };
        var comparer = StringComparer.CurrentCulture;
        var item = new DictionaryMapper<object>(mapping, comparer);

        Dictionary<string, object> itemMapping = Assert.IsType<Dictionary<string, object>>(item.MappingDictionary);
        Assert.Equal(mapping, itemMapping);
        Assert.Same(comparer, itemMapping.Comparer);
    }

    [Fact]
    public void Ctor_NullDictionary_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("mappingDictionary", () => new DictionaryMapper<int>(null!, StringComparer.CurrentCulture));
    }

    [Theory]
    [InlineData(null, false, CellMapperResult.HandleAction.IgnoreResultAndContinueMapping, null)]
    [InlineData("key", true, CellMapperResult.HandleAction.UseResultAndStopMapping, "value")]
    [InlineData("key2", true, CellMapperResult.HandleAction.UseResultAndStopMapping, 10)]
    [InlineData("no_such_key", false, CellMapperResult.HandleAction.IgnoreResultAndContinueMapping, null)]
    public void MapCellValue_ValidStringValue_ReturnsSuccess(string? stringValue, bool expectedSucceeded, CellMapperResult.HandleAction expectedAction, object? expectedValue)
    {
        var mapping = new Dictionary<string, object> { { "key", "value" }, { "KEY2", 10 } };
        var comparer = StringComparer.OrdinalIgnoreCase;
        var item = new DictionaryMapper<object>(mapping, comparer);

        CellMapperResult result = item.MapCellValue(new ReadCellResult(-1, stringValue));
        Assert.Equal(expectedSucceeded, result.Succeeded);
        Assert.Equal(expectedAction, result.Action);
        Assert.Equal(expectedValue, result.Value);
        Assert.Null(result.Exception);
    }
}

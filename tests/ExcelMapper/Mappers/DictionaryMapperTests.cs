using System;
using System.Collections.Generic;
using ExcelMapper.Abstractions;
using Xunit;

namespace ExcelMapper.Mappers.Tests
{
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
            Assert.Throws<ArgumentNullException>("mappingDictionary", () => new DictionaryMapper<int>(null, StringComparer.CurrentCulture));
        }

        [Theory]
        [InlineData("key", true, CellValueMapperResult.HandleAction.UseResultAndStopMapping, "value")]
        [InlineData("key2", true, CellValueMapperResult.HandleAction.UseResultAndStopMapping, 10)]
        [InlineData("no_such_key", false, CellValueMapperResult.HandleAction.IgnoreResultAndContinueMapping, null)]
        public void GetProperty_ValidStringValue_ReturnsSuccess(string stringValue, bool expectedSucceeded, CellValueMapperResult.HandleAction expectedAction, object expectedValue)
        {
            var mapping = new Dictionary<string, object> { { "key", "value" }, { "KEY2", 10 } };
            var comparer = StringComparer.OrdinalIgnoreCase;
            var item = new DictionaryMapper<object>(mapping, comparer);

            CellValueMapperResult result = item.MapCellValue(new ReadCellValueResult(-1, stringValue));
            Assert.Equal(expectedSucceeded, result.Succeeded);
            Assert.Equal(expectedAction, result.Action);
            Assert.Equal(expectedValue, result.Value);
            Assert.Null(result.Exception);
        }
    }
}

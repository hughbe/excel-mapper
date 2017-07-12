using System;
using System.Collections.Generic;
using Xunit;

namespace ExcelMapper.Mappings.Items.Tests
{
    public class MapStringValueMappingItemTests
    {
        [Fact]
        public void Ctor_Dictionary_ReturnsExpected()
        {
            var mapping = new Dictionary<string, object> { { "key", "value" } };
            var comparer = StringComparer.CurrentCulture;
            var item = new MapStringValueMappingItem<object>(mapping, comparer);

            Dictionary<string, object> itemMapping = Assert.IsType<Dictionary<string, object>>(item.MappingDictionary);
            Assert.Equal(mapping, itemMapping);
            Assert.Same(comparer, itemMapping.Comparer);
        }

        [Fact]
        public void Ctor_NullDictionary_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>("mappingDictionary", () => new MapStringValueMappingItem<int>(null, StringComparer.CurrentCulture));
        }

        [Theory]
        [InlineData("key", PropertyMappingResultType.Success, "value")]
        [InlineData("key2", PropertyMappingResultType.Success, 10)]
        [InlineData("no_such_key", PropertyMappingResultType.Continue, null)]
        public void GetProperty_ValidStringValue_ReturnsSuccess(string stringValue, PropertyMappingResultType expectedType, object expectedValue)
        {
            var mapping = new Dictionary<string, object> { { "key", "value" }, { "KEY2", 10 } };
            var comparer = StringComparer.OrdinalIgnoreCase;
            var item = new MapStringValueMappingItem<object>(mapping, comparer);

            PropertyMappingResult result = item.GetProperty(new ReadResult(-1, stringValue));
            Assert.Equal(expectedType, result.Type);
            Assert.Equal(expectedValue, result.Value);
        }
    }
}

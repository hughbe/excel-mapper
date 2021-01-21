﻿using System;
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
        [InlineData("key", PropertyMapperResultType.Success, "value")]
        [InlineData("key2", PropertyMapperResultType.Success, 10)]
        [InlineData("no_such_key", PropertyMapperResultType.Continue, null)]
        public void GetProperty_ValidStringValue_ReturnsSuccess(string stringValue, PropertyMapperResultType expectedType, object expectedValue)
        {
            var mapping = new Dictionary<string, object> { { "key", "value" }, { "KEY2", 10 } };
            var comparer = StringComparer.OrdinalIgnoreCase;
            var item = new DictionaryMapper<object>(mapping, comparer);

            object value = null;
            PropertyMapperResultType result = item.MapCellValue(new ReadCellValueResult(-1, -1, stringValue), ref value);
            Assert.Equal(expectedType, result);
            Assert.Equal(expectedValue, value);
        }
    }
}

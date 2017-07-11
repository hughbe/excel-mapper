using System;
using System.Collections.Generic;
using Xunit;

namespace ExcelMapper.Mappings.Items.Tests
{
    public class ChangeTypeMappingItemTests
    {
        [Theory]
        [InlineData(typeof(string))]
        [InlineData(typeof(int))]
        public void Ctor_Type(Type type)
        {
            var item = new ChangeTypeMappingItem(type);
            Assert.Equal(type, item.Type);
        }

        [Fact]
        public void Ctor_NullType_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>("type", () => new ChangeTypeMappingItem(null));
        }

        [Fact]
        public void Ctor_TypeNotIConvertible_ThrowsArgumentException()
        {
            Assert.Throws<ArgumentException>("type", () => new ChangeTypeMappingItem(typeof(List<int>)));
        }

        [Theory]
        [InlineData(typeof(int), "1", 1)]
        public void GetProperty_ValidStringValue_ReturnsSuccess(Type type, string stringValue, object expected)
        {
            var item = new ChangeTypeMappingItem(type);

            PropertyMappingResult result = item.GetProperty(null, 0, null, new MapResult(-1, stringValue));
            Assert.Equal(PropertyMappingResultType.Success, result.Type);
            Assert.Equal(expected, result.Value);
        }

        [Theory]
        [InlineData(typeof(uint), "-1")]
        [InlineData(typeof(uint), "abc")]
        [InlineData(typeof(uint), "")]
        [InlineData(typeof(uint), null)]
        public void GetProperty_InvalidStringValue_ReturnsInvalid(Type type, string stringValue)
        {
            var item = new ChangeTypeMappingItem(type);

            PropertyMappingResult result = item.GetProperty(null, 0, null, new MapResult(-1, stringValue));
            Assert.Equal(PropertyMappingResultType.Invalid, result.Type);
            Assert.Null(result.Value);
        }
    }
}

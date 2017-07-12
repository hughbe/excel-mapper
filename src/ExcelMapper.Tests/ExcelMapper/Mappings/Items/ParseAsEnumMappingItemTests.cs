using System;
using Xunit;

namespace ExcelMapper.Mappings.Items.Tests
{
    public class ParseAsEnumTests
    {
        [Theory]
        [InlineData(typeof(ConsoleColor))]
        public void Ctor_Type(Type enumType)
        {
            var item = new ParseAsEnumMappingItem(enumType);
            Assert.Same(enumType, item.EnumType);
        }

        [Fact]
        public void Ctor_NullEnumType_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>("enumType", () => new ParseAsEnumMappingItem(null));
        }

        [Theory]
        [InlineData(typeof(int))]
        [InlineData(typeof(Enum))]
        public void Ctor_EnumTypeNotEnum_ThrowsArgumentException(Type enumType)
        {
            Assert.Throws<ArgumentException>("enumType", () => new ParseAsEnumMappingItem(enumType));
        }

        [Theory]
        [InlineData(typeof(ConsoleColor), "Black", ConsoleColor.Black)]
        public void GetProperty_ValidStringValue_ReturnsSuccess(Type enumType, string stringValue, Enum expected)
        {
            var item = new ParseAsEnumMappingItem(enumType);

            PropertyMappingResult result = item.GetProperty(null, 0, null, new ReadResult(-1, stringValue));
            Assert.Equal(PropertyMappingResultType.Success, result.Type);
            Assert.Equal(expected, result.Value);
        }

        [Theory]
        [InlineData(typeof(ConsoleColor), null)]
        [InlineData(typeof(ConsoleColor), "")]
        [InlineData(typeof(ConsoleColor), "Invalid")]
        public void GetProperty_InvalidStringValue_ReturnsInvalid(Type enumType, string stringValue)
        {
            var item = new ParseAsEnumMappingItem(enumType);

            PropertyMappingResult result = item.GetProperty(null, 0, null, new ReadResult(-1, stringValue));
            Assert.Equal(PropertyMappingResultType.Invalid, result.Type);
            Assert.Null(result.Value);
        }
    }
}

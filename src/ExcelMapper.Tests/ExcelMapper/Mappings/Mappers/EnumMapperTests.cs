using System;
using Xunit;

namespace ExcelMapper.Mappings.Mappers.Tests
{
    public class EnumMapperTests
    {
        [Theory]
        [InlineData(typeof(ConsoleColor))]
        public void Ctor_Type(Type enumType)
        {
            var item = new EnumMapper(enumType);
            Assert.Same(enumType, item.EnumType);
        }

        [Fact]
        public void Ctor_NullEnumType_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>("enumType", () => new EnumMapper(null));
        }

        [Theory]
        [InlineData(typeof(int))]
        [InlineData(typeof(Enum))]
        public void Ctor_EnumTypeNotEnum_ThrowsArgumentException(Type enumType)
        {
            Assert.Throws<ArgumentException>("enumType", () => new EnumMapper(enumType));
        }

        [Theory]
        [InlineData(typeof(ConsoleColor), "Black", ConsoleColor.Black)]
        public void GetProperty_ValidStringValue_ReturnsSuccess(Type enumType, string stringValue, Enum expected)
        {
            var item = new EnumMapper(enumType);

            object value = null;
            PropertyMappingResultType result = item.GetProperty(new ReadCellValueResult(-1, stringValue), ref value);
            Assert.Equal(PropertyMappingResultType.Success, result);
            Assert.Equal(expected, value);
        }

        [Theory]
        [InlineData(typeof(ConsoleColor), null)]
        [InlineData(typeof(ConsoleColor), "")]
        [InlineData(typeof(ConsoleColor), "Invalid")]
        public void GetProperty_InvalidStringValue_ReturnsInvalid(Type enumType, string stringValue)
        {
            var item = new EnumMapper(enumType);

            object value = 1;
            PropertyMappingResultType result = item.GetProperty(new ReadCellValueResult(-1, stringValue), ref value);
            Assert.Equal(PropertyMappingResultType.Invalid, result);
            Assert.Equal(1, value);
        }
    }
}

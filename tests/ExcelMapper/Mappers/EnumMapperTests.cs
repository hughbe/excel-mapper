using System;
using System.Collections.Generic;
using ExcelMapper.Abstractions;
using Xunit;

namespace ExcelMapper.Mappers.Tests
{
    public class EnumMapperTests
    {
        [Theory]
        [InlineData(typeof(ConsoleColor))]
        public void Ctor_Type(Type enumType)
        {
            var item = new EnumMapper(enumType);
            Assert.Same(enumType, item.EnumType);
            Assert.False(item.IgnoreCase);
        }

        [Theory]
        [InlineData(typeof(ConsoleColor), true)]
        [InlineData(typeof(ConsoleColor), false)]
        public void Ctor_Type_Bool(Type enumType, bool ignoreCase)
        {
            var item = new EnumMapper(enumType, ignoreCase);
            Assert.Same(enumType, item.EnumType);
            Assert.Equal(ignoreCase, item.IgnoreCase);
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

        public static IEnumerable<object[]> GetProperty_ValidStringValue_TestData()
        {
            yield return new object[] { new EnumMapper(typeof(ConsoleColor)), "Black", ConsoleColor.Black };
            yield return new object[] { new EnumMapper(typeof(ConsoleColor), ignoreCase: true), "Black", ConsoleColor.Black };
            yield return new object[] { new EnumMapper(typeof(ConsoleColor), ignoreCase: true), "bLaCk", ConsoleColor.Black };
        }

        [Theory]
        [MemberData(nameof(GetProperty_ValidStringValue_TestData))]
        public void GetProperty_ValidStringValue_ReturnsSuccess(EnumMapper item, string stringValue, Enum expected)
        {
            
            CellValueMapperResult result = item.MapCellValue(new ReadCellValueResult(-1, stringValue));
            Assert.True(result.Succeeded);
            Assert.Equal(expected, result.Value);
            Assert.Null(result.Exception);
        }

        public static IEnumerable<object[]> GetProperty_InvalidStringValue_TestData()
        {
            yield return new object[] { new EnumMapper(typeof(ConsoleColor)), null };
            yield return new object[] { new EnumMapper(typeof(ConsoleColor)), "" };
            yield return new object[] { new EnumMapper(typeof(ConsoleColor)), "Invalid" };
            yield return new object[] { new EnumMapper(typeof(ConsoleColor)), "black" };
        }

        [Theory]
        [MemberData(nameof(GetProperty_InvalidStringValue_TestData))]
        public void GetProperty_InvalidStringValue_ReturnsInvalid(EnumMapper item, string stringValue)
        {
            CellValueMapperResult result = item.MapCellValue(new ReadCellValueResult(-1, stringValue));
            Assert.False(result.Succeeded);
            Assert.Null(result.Value);
            Assert.NotNull(result.Exception);
        }
    }
}

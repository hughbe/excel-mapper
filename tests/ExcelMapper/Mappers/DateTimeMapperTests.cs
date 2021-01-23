using System;
using System.Collections.Generic;
using System.Globalization;
using ExcelMapper.Abstractions;
using Xunit;

namespace ExcelMapper.Mappers.Tests
{
    public class DateTimeMapperTests
    {
        [Fact]
        public void Ctor_Default()
        {
            var item = new DateTimeMapper();
            Assert.Equal(new string[] { "G" }, item.Formats);
            Assert.Null(item.Provider);
            Assert.Equal(DateTimeStyles.None, item.Style);
        }

        [Fact]
        public void Formats_SetValid_GetReturnsExpected()
        {
            var formats = new string[] { null, "", "abc" };
            var item = new DateTimeMapper { Formats = formats };
            Assert.Same(formats, item.Formats);
        }

        [Fact]
        public void Formats_SetNull_ThrowsArgumentNullException()
        {
            var item = new DateTimeMapper();
            Assert.Throws<ArgumentNullException>("value", () => item.Formats = null);
        }

        [Fact]
        public void Formats_SetEmpty_ThrowsArgumentException()
        {
            var item = new DateTimeMapper();
            Assert.Throws<ArgumentException>("value", () => item.Formats = new string[0]);
        }

        [Fact]
        public void Provider_Set_GetReturnsExpected()
        {
            IFormatProvider provider = CultureInfo.CurrentCulture;
            var item = new DateTimeMapper { Provider = provider};
            Assert.Same(provider, item.Provider);
        }

        [Theory]
        [InlineData(DateTimeStyles.AdjustToUniversal)]
        [InlineData((DateTimeStyles)int.MaxValue)]
        public void Styles_Set_GetReturnsExpected(DateTimeStyles style)
        {
            var item = new DateTimeMapper { Style = style };
            Assert.Equal(style, item.Style);
        }

        public static IEnumerable<object[]> GetProperty_Valid_TestData()
        {
            yield return new object[] { new DateTime(2017, 7, 12, 7, 57, 46).ToString("G"), new string[] { "G" }, DateTimeStyles.None, new DateTime(2017, 7, 12, 7, 57, 46) };
            yield return new object[] { new DateTime(2017, 7, 12, 7, 57, 46).ToString("G"), new string[] { "G", "yyyy-MM-dd" }, DateTimeStyles.None, new DateTime(2017, 7, 12, 7, 57, 46) };
            yield return new object[] { "   2017-07-12   ", new string[] { "G", "yyyy-MM-dd" }, DateTimeStyles.AllowWhiteSpaces, new DateTime(2017, 7, 12) };
        }

        [Theory]
        [MemberData(nameof(GetProperty_Valid_TestData))]
        public void GetProperty_ValidStringValue_ReturnsSuccess(string stringValue, string[] formats, DateTimeStyles style, DateTime expected)
        {
            var item = new DateTimeMapper
            {
                Formats = formats,
                Style = style
            };

            CellValueMapperResult result = item.MapCellValue(new ReadCellValueResult(-1, stringValue));
            Assert.True(result.Succeeded);
            Assert.Equal(expected, result.Value);
            Assert.Null(result.Exception);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("invalid")]
        [InlineData("12/07/2017 07:57:61")]
        public void GetProperty_InvalidStringValue_ReturnsInvalid(string stringValue)
        {
            var item = new DateTimeMapper();
            CellValueMapperResult result = item.MapCellValue(new ReadCellValueResult(-1, stringValue));
            Assert.False(result.Succeeded);
            Assert.Null(result.Value);
            Assert.NotNull(result.Exception);
        }
    }
}

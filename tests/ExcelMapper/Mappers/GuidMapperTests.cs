using System;
using System.Collections.Generic;
using ExcelMapper.Abstractions;
using Xunit;

namespace ExcelMapper.Mappers.Tests
{
    public class GuidMapperTests
    {
        public static IEnumerable<object[]> GetProperty_ValidStringValue_TestData()
        {
            yield return new object[] { "a8a110d5fc4943c5bf46802db8f843ff", new Guid("a8a110d5fc4943c5bf46802db8f843ff") };
            yield return new object[] { "a8a110d5-fc49-43c5-bf46-802db8f843ff", new Guid("a8a110d5fc4943c5bf46802db8f843ff") };
            yield return new object[] { "{a8a110d5-fc49-43c5-bf46-802db8f843ff}", new Guid("a8a110d5fc4943c5bf46802db8f843ff") };
            yield return new object[] { "(a8a110d5-fc49-43c5-bf46-802db8f843ff)", new Guid("a8a110d5fc4943c5bf46802db8f843ff") };
            yield return new object[] { "{0xa8a110d5,0xfc49,0x43c5,{0xbf,0x46,0x80,0x2d,0xb8,0xf8,0x43,0xff}}", new Guid("a8a110d5fc4943c5bf46802db8f843ff") };
        }

        [Theory]
        [MemberData(nameof(GetProperty_ValidStringValue_TestData))]
        public void GetProperty_ValidStringValue_ReturnsSuccess(string stringValue, Guid expected)
        {
            var item = new GuidMapper();

            CellValueMapperResult result = item.MapCellValue(new ReadCellValueResult(-1, stringValue));
            Assert.True(result.Succeeded);
            Assert.Equal(expected, result.Value);
            Assert.Null(result.Exception);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("invalid")]
        public void GetProperty_InvalidStringValue_ReturnsInvalid(string stringValue)
        {
            var item = new GuidMapper();

            CellValueMapperResult result = item.MapCellValue(new ReadCellValueResult(-1, stringValue));
            Assert.False(result.Succeeded);
            Assert.Null(result.Value);
            Assert.NotNull(result.Exception);
        }
    }
}

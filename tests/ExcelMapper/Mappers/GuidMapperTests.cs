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

            object value = null;
            PropertyMapperResultType result = item.MapCellValue(new ReadCellValueResult(-1, stringValue), ref value);
            Assert.Equal(PropertyMapperResultType.Success, result);
            Assert.Equal(expected, value);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("invalid")]
        public void GetProperty_InvalidStringValue_ReturnsInvalid(string stringValue)
        {
            var item = new GuidMapper();

            object value = 1;
            PropertyMapperResultType result = item.MapCellValue(new ReadCellValueResult(-1, stringValue), ref value);
            Assert.Equal(PropertyMapperResultType.Invalid, result);
            Assert.Equal(1, value);
        }
    }
}

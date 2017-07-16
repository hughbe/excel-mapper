using System;
using System.Collections.Generic;
using Xunit;

namespace ExcelMapper.Mappings.Mappers.Tests
{
    public class UriMapperTests
    {
        public static IEnumerable<object[]> GetProperty_TestData()
        {
            yield return new object[] { "http://microsoft.com", new Uri("http://microsoft.com") };
        }

        [Theory]
        [MemberData(nameof(GetProperty_TestData))]
        public void GetProperty_ValidStringValue_ReturnsSuccess(string stringValue, Uri expected)
        {
            var item = new UriMapper();

            object value = null;
            PropertyMappingResultType result = item.GetProperty(new ReadResult(-1, stringValue), ref value);
            Assert.Equal(PropertyMappingResultType.Success, result);
            Assert.Equal(expected, value);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("invalid")]
        [InlineData("/relative")]
        public void GetProperty_InvalidStringValue_ReturnsInvalid(string stringValue)
        {
            var item = new UriMapper();

            object value = 1;
            PropertyMappingResultType result = item.GetProperty(new ReadResult(-1, stringValue), ref value);
            Assert.Equal(PropertyMappingResultType.Invalid, result);
            Assert.Equal(1, value);
        }
    }
}

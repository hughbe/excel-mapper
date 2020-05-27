using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
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
            var item = new UriMapper();

            object value = 1;
            PropertyMapperResultType result = item.MapCellValue(new ReadCellValueResult(-1, stringValue), ref value);
            Assert.Equal(PropertyMapperResultType.Invalid, result);
            Assert.Equal(1, value);
        }

        [Theory]
        [InlineData("/relative")]
        public void GetProperty_InvalidStringValueWindows_ReturnsInvalid(string stringValue)
        {
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                return;
            }

            GetProperty_InvalidStringValue_ReturnsInvalid(stringValue);
        }
    }
}

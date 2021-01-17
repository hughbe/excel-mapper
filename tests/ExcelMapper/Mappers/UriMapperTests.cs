using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using ExcelMapper.Abstractions;
using Xunit;

namespace ExcelMapper.Mappers.Tests
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
            var item = new UriMapper();
            CellValueMapperResult result = item.MapCellValue(new ReadCellValueResult(-1, stringValue));
            Assert.False(result.Succeeded);
            Assert.Null(result.Value);
            Assert.NotNull(result.Exception);
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

using Xunit;

namespace ExcelMapper.Mappings.Items.Tests
{
    public class ParseAsBoolMappingItemTestsTests
    {
        [Theory]
        [InlineData("1", true)]
        [InlineData("0", false)]
        [InlineData("true", true)]
        [InlineData("false", false)]
        public void GetProperty_ValidStringValue_ReturnsSuccess(string stringValue, bool expected)
        {
            var item = new ParseAsBoolMappingItem();

            PropertyMappingResult result = item.GetProperty(new ReadResult(-1, stringValue));
            Assert.Equal(PropertyMappingResultType.Success, result.Type);
            Assert.Equal(expected, result.Value);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("invalid")]
        public void GetProperty_InvalidStringValue_ReturnsInvalid(string stringValue)
        {
            var item = new ParseAsBoolMappingItem();

            PropertyMappingResult result = item.GetProperty(new ReadResult(-1, stringValue));
            Assert.Equal(PropertyMappingResultType.Invalid, result.Type);
            Assert.Null(result.Value);
        }
    }
}

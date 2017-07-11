using Xunit;

namespace ExcelMapper.Mappings.Items.Tests
{
    public class ParseAsStringMappingItemTests
    {
        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("abc")]
        public void GetProperty_Invoke_ReturnsBegan(string stringValue)
        {
            var item = new ParseAsStringMappingItem();

            PropertyMappingResult result = item.GetProperty(null, 0, null, new MapResult(-1, stringValue));
            Assert.Equal(PropertyMappingResultType.Began, result.Type);
            Assert.Same(stringValue, result.Value);
        }
    }
}

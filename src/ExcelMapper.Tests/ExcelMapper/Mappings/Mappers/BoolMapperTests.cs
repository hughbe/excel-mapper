using Xunit;

namespace ExcelMapper.Mappings.Mappers.Tests
{
    public class BoolMapperTests
    {
        [Theory]
        [InlineData("1", true)]
        [InlineData("0", false)]
        [InlineData("true", true)]
        [InlineData("false", false)]
        public void GetProperty_ValidStringValue_ReturnsSuccess(string stringValue, bool expected)
        {
            var item = new BoolMapper();

            object value = null;
            PropertyMappingResultType result = item.GetProperty(new ReadResult(-1, stringValue), ref value);
            Assert.Equal(PropertyMappingResultType.Success, result);
            Assert.Equal(expected, value);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("invalid")]
        public void GetProperty_InvalidStringValue_ReturnsInvalid(string stringValue)
        {
            var item = new BoolMapper();

            object value = 1;
            PropertyMappingResultType result = item.GetProperty(new ReadResult(-1, stringValue), ref value);
            Assert.Equal(PropertyMappingResultType.Invalid, result);
            Assert.Equal(1, value);
        }
    }
}

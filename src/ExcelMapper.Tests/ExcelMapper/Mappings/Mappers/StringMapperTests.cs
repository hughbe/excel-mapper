using Xunit;

namespace ExcelMapper.Mappings.Mappers.Tests
{
    public class StringMapperTests
    {
        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("abc")]
        public void GetProperty_Invoke_ReturnsBegan(string stringValue)
        {
            var item = new StringMapper();

            object value = 1;
            PropertyMapperResultType result = item.GetProperty(new ReadCellValueResult(-1, stringValue), ref value);
            Assert.Equal(PropertyMapperResultType.SuccessIfNoOtherSuccess, result);
            Assert.Same(stringValue, value);
        }
    }
}

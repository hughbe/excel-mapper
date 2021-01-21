using ExcelMapper.Abstractions;
using Xunit;

namespace ExcelMapper.Mappers.Tests
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
            PropertyMapperResultType result = item.MapCellValue(new ReadCellValueResult(-1, -1, stringValue), ref value);
            Assert.Equal(PropertyMapperResultType.SuccessIfNoOtherSuccess, result);
            Assert.Same(stringValue, value);
        }
    }
}

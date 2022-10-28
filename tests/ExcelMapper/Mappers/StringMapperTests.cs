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

            CellValueMapperResult result = item.MapCell(new ExcelCell(null, -1, -1), new CellValueMapperResult(stringValue, null, CellValueMapperResult.HandleAction.UseResultAndStopMapping), null);
            Assert.True(result.Succeeded);
            Assert.Equal(stringValue, result.Value);
            Assert.Null(result.Exception);
        }
    }
}

using ExcelMapper.Abstractions;
using Xunit;

namespace ExcelMapper.Transformers.Tests
{
    public class TrimCellValueTransformerTests
    {
        [Theory]
        [InlineData(null, null)]
        [InlineData("", "")]
        [InlineData(" abc ", "abc")]
        public void TransformStringValue_Invoke_ReturnsExpected(string stringValue, string expected)
        {
            var transformer = new TrimCellValueTransformer();
            Assert.Equal(expected, transformer.TransformStringValue(null, 0, new ReadCellValueResult(-1, stringValue)));
        }
    }
}

using Xunit;

namespace ExcelMapper.Mappings.Transformers.Tests
{
    public class TrimStringTransformerTests
    {
        [Theory]
        [InlineData(null, null)]
        [InlineData("", "")]
        [InlineData(" abc ", "abc")]
        public void TransformStringValue_Invoke_ReturnsExpected(string stringValue, string expected)
        {
            var transformer = new TrimStringTransformer();
            Assert.Equal(expected, transformer.TransformStringValue(null, 0, null, new MapResult(-1, stringValue)));
        }
    }
}

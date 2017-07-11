using Xunit;

namespace ExcelMapper.Mappings.Tests
{
    public class PropertyMappingResultTests
    {
        [Fact]
        public void Ctor_Default()
        {
            var result = new PropertyMappingResult();
            Assert.Equal(PropertyMappingResultType.Began, result.Type);
            Assert.Null(result.Value);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("value")]
        public void Began_Invoke_ReturnsExpected(object value)
        {
            var result = PropertyMappingResult.Began(value);
            Assert.Equal(PropertyMappingResultType.Began, result.Type);
            Assert.Same(value, result.Value);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("value")]
        public void Success_Invoke_ReturnsExpected(object value)
        {
            var result = PropertyMappingResult.Success(value);
            Assert.Equal(PropertyMappingResultType.Success, result.Type);
            Assert.Same(value, result.Value);
        }

        [Fact]
        public void Continue_Invoke_ReturnsExpected()
        {
            var result = PropertyMappingResult.Continue();
            Assert.Equal(PropertyMappingResultType.Continue, result.Type);
            Assert.Null(result.Value);
        }

        [Fact]
        public void Invalid_Invoke_ReturnsExpected()
        {
            var result = PropertyMappingResult.Invalid();
            Assert.Equal(PropertyMappingResultType.Invalid, result.Type);
            Assert.Null(result.Value);
        }
    }
}

using Xunit;

namespace ExcelMapper.Mappings.Fallbacks.Tests
{
    public class FixedValueFallbackTests
    {
        [Theory]
        [InlineData(null)]
        [InlineData(1)]
        [InlineData("value")]
        public void Ctor_Default(object value)
        {
            var fallback = new FixedValueFallback(value);
            Assert.Same(value, value);

            PropertyMappingResult result = fallback.GetProperty(null, 0, null, new ReadResult());
            Assert.Equal(PropertyMappingResultType.Success, result.Type);
            Assert.Same(value, result.Value);
        }
    }
}

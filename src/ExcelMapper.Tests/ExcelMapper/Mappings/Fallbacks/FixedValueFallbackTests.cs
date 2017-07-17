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

            object result = fallback.PerformFallback(null, 0, new ReadCellValueResult());
            Assert.Same(value, result);
        }
    }
}

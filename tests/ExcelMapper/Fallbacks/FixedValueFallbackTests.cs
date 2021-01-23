using ExcelMapper.Abstractions;
using Xunit;

namespace ExcelMapper.Fallbacks.Tests
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
            Assert.Same(value, fallback.Value);

            object result = fallback.PerformFallback(null, 0, new ReadCellValueResult(), null, null);
            Assert.Same(value, result);
        }
    }
}

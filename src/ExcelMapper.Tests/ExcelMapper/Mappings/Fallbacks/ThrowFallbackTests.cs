using Xunit;

namespace ExcelMapper.Mappings.Fallbacks.Tests
{
    public class ThrowFallbackTests
    {
        [Fact]
        public void GetProperty_Invoke_ThrowsExcelMappingException()
        {
            var fallback = new ThrowFallback();
            Assert.Throws<ExcelMappingException>(() => fallback.GetProperty(null, 0, null, new ReadResult()));
        }
    }
}

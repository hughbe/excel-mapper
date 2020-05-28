using Xunit;

namespace ExcelMapper.Tests
{
    public class ExcelIgnoreAttributeTests
    {
        [Fact]
        public void Ctor_Default()
        {
            new ExcelIgnoreAttribute();
        }
    }
}

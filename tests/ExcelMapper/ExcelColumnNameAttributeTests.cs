using System;
using Xunit;

namespace ExcelMapper.Tests
{
    public class ExcelColumnNameAttributeTests
    {
        [Theory]
        [InlineData("columnname")]
        [InlineData("ColumnName")]
        public void Ctor_String(string name)
        {
            var attribute = new ExcelColumnNameAttribute(name);
            Assert.Equal(name, attribute.Name);
        }

        [Fact]
        public void Ctor_NullName_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>("name", () => new ExcelColumnNameAttribute(null));
        }

        [Fact]
        public void Ctor_EmptyName_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentException>("name", () => new ExcelColumnNameAttribute(string.Empty));
        }
    }
}

using System;
using Xunit;

namespace ExcelMapper.Mappings.Tests
{
    public class ColumnPropertyMapperTests
    {
        [Fact]
        public void Ctor_ColumnName()
        {
            var mapper = new ColumnPropertyMapper("ColumnName");
            Assert.Equal("ColumnName", mapper.ColumnName);
        }

        [Fact]
        public void Ctor_NullColumnName_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>("columnName", () => new ColumnPropertyMapper(null));
        }

        [Fact]
        public void Ctor_EmptyColumnName_ThrowsArgumentException()
        {
            Assert.Throws<ArgumentException>("columnName", () => new ColumnPropertyMapper(string.Empty));
        }
    }
}

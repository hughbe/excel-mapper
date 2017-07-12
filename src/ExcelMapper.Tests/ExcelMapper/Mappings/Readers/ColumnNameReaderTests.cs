using System;
using Xunit;

namespace ExcelMapper.Mappings.Readers.Tests
{
    public class ColumnNameReaderTests
    {
        [Fact]
        public void Ctor_ColumnName()
        {
            var reader = new ColumnNameReader("ColumnName");
            Assert.Equal("ColumnName", reader.ColumnName);
        }

        [Fact]
        public void Ctor_NullColumnName_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>("columnName", () => new ColumnNameReader(null));
        }

        [Fact]
        public void Ctor_EmptyColumnName_ThrowsArgumentException()
        {
            Assert.Throws<ArgumentException>("columnName", () => new ColumnNameReader(string.Empty));
        }
    }
}

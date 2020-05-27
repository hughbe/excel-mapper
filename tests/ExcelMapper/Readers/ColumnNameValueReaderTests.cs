using System;
using Xunit;

namespace ExcelMapper.Readers.Tests
{
    public class ColumnNameValueReaderTests
    {
        [Fact]
        public void Ctor_ColumnName()
        {
            var reader = new ColumnNameValueReader("ColumnName");
            Assert.Equal("ColumnName", reader.ColumnName);
        }

        [Fact]
        public void Ctor_NullColumnName_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>("columnName", () => new ColumnNameValueReader(null));
        }

        [Fact]
        public void Ctor_EmptyColumnName_ThrowsArgumentException()
        {
            Assert.Throws<ArgumentException>("columnName", () => new ColumnNameValueReader(string.Empty));
        }
    }
}

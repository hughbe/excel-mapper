using System;
using Xunit;

namespace ExcelMapper.Readers.Tests
{
    public class ColumnIndexValueReaderTests
    {
        [Theory]
        [InlineData(0)]
        [InlineData(10)]
        public void Ctor_ColumnIndex(int columnIndex)
        {
            var reader = new ColumnIndexValueReader(columnIndex);
            Assert.Equal(columnIndex, reader.ColumnIndex);
        }

        [Fact]
        public void Ctor_NegativeColumnIndex_ThrowsArgumentOutOfRangeException()
        {
            Assert.Throws<ArgumentOutOfRangeException>("columnIndex", () => new ColumnIndexValueReader(-1));
        }
    }
}

using System;
using Xunit;

namespace ExcelMapper.Mappings.Readers.Tests
{
    public class ColumnIndexReaderTests
    {
        [Theory]
        [InlineData(0)]
        [InlineData(10)]
        public void Ctor_ColumnIndex(int columnIndex)
        {
            var reader = new ColumnIndexReader(columnIndex);
            Assert.Equal(columnIndex, reader.ColumnIndex);
        }

        [Fact]
        public void Ctor_NegativeColumnIndex_ThrowsArgumentOutOfRangeException()
        {
            Assert.Throws<ArgumentOutOfRangeException>("columnIndex", () => new ColumnIndexReader(-1));
        }
    }
}

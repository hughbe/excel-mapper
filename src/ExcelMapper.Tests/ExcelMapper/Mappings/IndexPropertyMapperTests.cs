using System;
using Xunit;

namespace ExcelMapper.Mappings.Tests
{
    public class ColumnPropertyMapperTetss
    {
        [Theory]
        [InlineData(0)]
        [InlineData(10)]
        public void Ctor_ColumnIndex(int columnIndex)
        {
            var mapper = new IndexPropertyMapper(columnIndex);
            Assert.Equal(columnIndex, mapper.ColumnIndex);
        }

        [Fact]
        public void Ctor_NegativeColumnIndex_ThrowsArgumentOutOfRangeException()
        {
            Assert.Throws<ArgumentOutOfRangeException>("columnIndex", () => new IndexPropertyMapper(-1));
        }
    }
}

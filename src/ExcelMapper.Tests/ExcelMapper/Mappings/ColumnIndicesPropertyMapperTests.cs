using System;
using Xunit;

namespace ExcelMapper.Mappings.Tests
{
    public class ColumnIndicesPropertyMapperTests
    {
        [Fact]
        public void Ctor_ColumnIndices()
        {
            var columnIndices = new int[] { 0, 1 };
            var mapper = new ColumnIndicesPropertyMapper(columnIndices);
            Assert.Equal(columnIndices, mapper.ColumnIndices);
        }

        [Fact]
        public void Ctor_NullColumnIndices_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>("columnIndices", () => new ColumnIndicesPropertyMapper(null));
        }

        [Fact]
        public void Ctor_EmptyColumnNames_ThrowsArgumentException()
        {
            Assert.Throws<ArgumentException>("columnIndices", () => new ColumnIndicesPropertyMapper(new int[0]));
        }

        [Fact]
        public void Ctor_NegativeValueInColumnIndices_ThrowsArgumentOutOfRangeException()
        {
            Assert.Throws<ArgumentOutOfRangeException>("columnIndices", () => new ColumnIndicesPropertyMapper(new int[] { -1 }));
        }
    }
}

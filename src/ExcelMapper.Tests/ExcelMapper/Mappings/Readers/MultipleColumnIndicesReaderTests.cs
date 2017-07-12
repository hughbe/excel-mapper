using System;
using Xunit;

namespace ExcelMapper.Mappings.Readers.Tests
{
    public class MultipleColumnIndicesReaderTests
    {
        [Fact]
        public void Ctor_ColumnIndices()
        {
            var columnIndices = new int[] { 0, 1 };
            var reader = new MultipleColumnIndicesReader(columnIndices);
            Assert.Same(columnIndices, reader.ColumnIndices);
        }

        [Fact]
        public void Ctor_NullColumnIndices_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>("columnIndices", () => new MultipleColumnIndicesReader(null));
        }

        [Fact]
        public void Ctor_EmptyColumnNames_ThrowsArgumentException()
        {
            Assert.Throws<ArgumentException>("columnIndices", () => new MultipleColumnIndicesReader(new int[0]));
        }

        [Fact]
        public void Ctor_NegativeValueInColumnIndices_ThrowsArgumentOutOfRangeException()
        {
            Assert.Throws<ArgumentOutOfRangeException>("columnIndices", () => new MultipleColumnIndicesReader(new int[] { -1 }));
        }
    }
}

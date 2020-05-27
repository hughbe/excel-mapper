using System;
using Xunit;

namespace ExcelMapper.Readers.Tests
{
    public class MultipleColumnIndicesValueReaderTests
    {
        [Fact]
        public void Ctor_ColumnIndices()
        {
            var columnIndices = new int[] { 0, 1 };
            var reader = new MultipleColumnIndicesValueReader(columnIndices);
            Assert.Same(columnIndices, reader.ColumnIndices);
        }

        [Fact]
        public void Ctor_NullColumnIndices_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>("columnIndices", () => new MultipleColumnIndicesValueReader(null));
        }

        [Fact]
        public void Ctor_EmptyColumnNames_ThrowsArgumentException()
        {
            Assert.Throws<ArgumentException>("columnIndices", () => new MultipleColumnIndicesValueReader(new int[0]));
        }

        [Fact]
        public void Ctor_NegativeValueInColumnIndices_ThrowsArgumentOutOfRangeException()
        {
            Assert.Throws<ArgumentOutOfRangeException>("columnIndices", () => new MultipleColumnIndicesValueReader(new int[] { -1 }));
        }
    }
}

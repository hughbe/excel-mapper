using System;
using Xunit;

namespace ExcelMapper.Mappings.Readers.Tests
{
    public class SplitColumnReaderTests
    {
        [Fact]
        public void Ctor_ColumnReader()
        {
            var innerReader = new ColumnNameReader("ColumnName");
            var reader = new SplitColumnReader(innerReader);
            Assert.Same(innerReader, reader.ColumnReader);

            Assert.Equal(StringSplitOptions.None, reader.Options);
            Assert.Equal(new char[] { ',' }, reader.Separators);
        }

        [Fact]
        public void Ctor_NullColumnReader_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>("columnReader", () => new SplitColumnReader(null));
        }

        [Theory]
        [InlineData(new char[] { ',' })]
        [InlineData(new char[] { ',', ';' })]
        public void Separators_SetValid_GetReturnsExpected(char[] separators)
        {
            var reader = new SplitColumnReader(new ColumnNameReader("ColumnName")) { Separators = separators };
            Assert.Same(separators, reader.Separators);
        }

        [Fact]
        public void Separators_SetNull_ThrowsArgumentNullException()
        {
            var reader = new SplitColumnReader(new ColumnNameReader("ColumnName"));
            Assert.Throws<ArgumentNullException>("value", () => reader.Separators = null);
        }

        [Fact]
        public void Separators_SetEmpty_ThrowsArgumentException()
        {
            var reader = new SplitColumnReader(new ColumnNameReader("ColumnName"));
            Assert.Throws<ArgumentException>("value", () => reader.Separators = new char[0]);
        }

        [Theory]
        [InlineData(StringSplitOptions.None - 1)]
        [InlineData(StringSplitOptions.None)]
        [InlineData(StringSplitOptions.RemoveEmptyEntries)]
        [InlineData(StringSplitOptions.RemoveEmptyEntries + 1)]
        public void Options_Set_GetReturnsExpected(StringSplitOptions options)
        {
            var reader = new SplitColumnReader(new ColumnNameReader("ColumnName")) { Options = options };
            Assert.Equal(options, reader.Options);
        }

        [Fact]
        public void ColumnReader_SetValid_GetReturnsExpected()
        {
            var innerReader = new ColumnNameReader("ColumnName1");
            var reader = new SplitColumnReader(new ColumnNameReader("ColumnName2")) { ColumnReader = innerReader };

            Assert.Same(innerReader, reader.ColumnReader);
        }

        [Fact]
        public void ColumnReader_SetNull_ThrowsArgumentNullException()
        {
            var reader = new SplitColumnReader(new ColumnNameReader("ColumnName"));
            Assert.Throws<ArgumentNullException>("value", () => reader.ColumnReader = null);
        }
    }
}

using System;
using ExcelDataReader;
using Xunit;

namespace ExcelMapper.Readers.Tests
{
    public class CharSplitCellValueReaderTests
    {
        [Fact]
        public void Ctor_CellReader()
        {
            var innerReader = new ColumnNameValueReader("ColumnName");
            var reader = new CharSplitCellValueReader(innerReader);
            Assert.Same(innerReader, reader.CellReader);

            Assert.Equal(StringSplitOptions.None, reader.Options);
            Assert.Equal(new char[] { ',' }, reader.Separators);
        }

        [Theory]
        [InlineData(new char[] { ',' })]
        [InlineData(new char[] { ',', ';' })]
        public void Separators_SetValid_GetReturnsExpected(char[] separators)
        {
            var reader = new CharSplitCellValueReader(new ColumnNameValueReader("ColumnName")) { Separators = separators };
            Assert.Same(separators, reader.Separators);
        }

        [Fact]
        public void Separators_SetNull_ThrowsArgumentNullException()
        {
            var reader = new CharSplitCellValueReader(new ColumnNameValueReader("ColumnName"));
            Assert.Throws<ArgumentNullException>("value", () => reader.Separators = null);
        }

        [Fact]
        public void Separators_SetEmpty_ThrowsArgumentException()
        {
            var reader = new CharSplitCellValueReader(new ColumnNameValueReader("ColumnName"));
            Assert.Throws<ArgumentException>("value", () => reader.Separators = new char[0]);
        }

        [Theory]
        [InlineData(StringSplitOptions.None - 1)]
        [InlineData(StringSplitOptions.None)]
        [InlineData(StringSplitOptions.RemoveEmptyEntries)]
        [InlineData(StringSplitOptions.RemoveEmptyEntries + 1)]
        public void Options_Set_GetReturnsExpected(StringSplitOptions options)
        {
            var reader = new CharSplitCellValueReader(new ColumnNameValueReader("ColumnName")) { Options = options };
            Assert.Equal(options, reader.Options);
        }
    }
}

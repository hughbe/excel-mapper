using System;
using ExcelDataReader;
using Xunit;

namespace ExcelMapper.Mappings.Readers.Tests
{
    public class SplitCellValueReaderTests
    {
        [Fact]
        public void Ctor_CellReader()
        {
            var innerReader = new ColumnNameValueReader("ColumnName");
            var reader = new SplitCellValueReader(innerReader);
            Assert.Same(innerReader, reader.CellReader);

            Assert.Equal(StringSplitOptions.None, reader.Options);
            Assert.Equal(new char[] { ',' }, reader.Separators);
        }

        [Fact]
        public void Ctor_NullCellReader_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>("cellReader", () => new SplitCellValueReader(null));
        }

        [Theory]
        [InlineData(new char[] { ',' })]
        [InlineData(new char[] { ',', ';' })]
        public void Separators_SetValid_GetReturnsExpected(char[] separators)
        {
            var reader = new SplitCellValueReader(new ColumnNameValueReader("ColumnName")) { Separators = separators };
            Assert.Same(separators, reader.Separators);
        }

        [Fact]
        public void Separators_SetNull_ThrowsArgumentNullException()
        {
            var reader = new SplitCellValueReader(new ColumnNameValueReader("ColumnName"));
            Assert.Throws<ArgumentNullException>("value", () => reader.Separators = null);
        }

        [Fact]
        public void Separators_SetEmpty_ThrowsArgumentException()
        {
            var reader = new SplitCellValueReader(new ColumnNameValueReader("ColumnName"));
            Assert.Throws<ArgumentException>("value", () => reader.Separators = new char[0]);
        }

        [Theory]
        [InlineData(StringSplitOptions.None - 1)]
        [InlineData(StringSplitOptions.None)]
        [InlineData(StringSplitOptions.RemoveEmptyEntries)]
        [InlineData(StringSplitOptions.RemoveEmptyEntries + 1)]
        public void Options_Set_GetReturnsExpected(StringSplitOptions options)
        {
            var reader = new SplitCellValueReader(new ColumnNameValueReader("ColumnName")) { Options = options };
            Assert.Equal(options, reader.Options);
        }

        [Fact]
        public void CellReader_SetValid_GetReturnsExpected()
        {
            var innerReader = new ColumnNameValueReader("ColumnName1");
            var reader = new SplitCellValueReader(new ColumnNameValueReader("ColumnName2")) { CellReader = innerReader };

            Assert.Same(innerReader, reader.CellReader);
        }

        [Fact]
        public void CellReader_SetNull_ThrowsArgumentNullException()
        {
            var reader = new SplitCellValueReader(new ColumnNameValueReader("ColumnName"));
            Assert.Throws<ArgumentNullException>("value", () => reader.CellReader = null);
        }

        [Fact]
        public void GetValues_NullReaderValue_ReturnsEmpty()
        {
            var reader = new SplitCellValueReader(new NullValueReader());
            Assert.Empty(reader.GetValues(null, 0, null));
        }

        private class NullValueReader : ICellValueReader
        {
            public ReadCellValueResult GetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader)
            {
                return new ReadCellValueResult();
            }
        }
    }
}

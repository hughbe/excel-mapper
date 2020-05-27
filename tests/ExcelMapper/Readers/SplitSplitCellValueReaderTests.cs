using System;
using System.Collections.Generic;
using ExcelDataReader;
using ExcelMapper.Abstractions;
using Xunit;

namespace ExcelMapper.Readers.Tests
{
    public class SplitCellValueReaderTests
    {
        [Fact]
        public void Ctor_CellReader()
        {
            var innerReader = new ColumnNameValueReader("ColumnName");
            var reader = new SubSplitCellValueReader(innerReader);
            Assert.Same(innerReader, reader.CellReader);

            Assert.Equal(StringSplitOptions.None, reader.Options);
        }

        [Fact]
        public void Ctor_NullCellReader_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>("cellReader", () => new SubSplitCellValueReader(null));
        }

        [Theory]
        [InlineData(StringSplitOptions.None - 1)]
        [InlineData(StringSplitOptions.None)]
        [InlineData(StringSplitOptions.RemoveEmptyEntries)]
        [InlineData(StringSplitOptions.RemoveEmptyEntries + 1)]
        public void Options_Set_GetReturnsExpected(StringSplitOptions options)
        {
            var reader = new SubSplitCellValueReader(new ColumnNameValueReader("ColumnName")) { Options = options };
            Assert.Equal(options, reader.Options);
        }

        [Fact]
        public void CellReader_SetValid_GetReturnsExpected()
        {
            var innerReader = new ColumnNameValueReader("ColumnName1");
            var reader = new SubSplitCellValueReader(new ColumnNameValueReader("ColumnName2")) { CellReader = innerReader };

            Assert.Same(innerReader, reader.CellReader);
        }

        [Fact]
        public void CellReader_SetNull_ThrowsArgumentNullException()
        {
            var reader = new SubSplitCellValueReader(new ColumnNameValueReader("ColumnName"));
            Assert.Throws<ArgumentNullException>("value", () => reader.CellReader = null);
        }

        [Fact]
        public void GetValues_NullReaderValue_ReturnsEmpty()
        {
            var reader = new SubSplitCellValueReader(new NullValueReader());
            Assert.True(reader.TryGetValues(null, 0, null, out IEnumerable<ReadCellValueResult> result));
            Assert.Empty(result);
        }

        private class SubSplitCellValueReader : SplitCellValueReader
        {
            public SubSplitCellValueReader(ISingleCellValueReader cellReader) : base(cellReader)
            {
            }

            protected override string[] GetValues(string value) => new string[0];
        }

        private class NullValueReader : ISingleCellValueReader
        {
            public bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, out ReadCellValueResult result)
            {
                result = new ReadCellValueResult();
                return true;
            }
        }
    }
}

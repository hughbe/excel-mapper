using Xunit;

namespace ExcelMapper.Abstractions.Tests
{
    public class ReadCellValueResultTests
    {
        [Fact]
        public void Ctor_Default()
        {
            var result = new ReadCellValueResult();
            Assert.Equal(0, result.ColumnIndex);
            Assert.Null(result.StringValue);
        }

        [Theory]
        [InlineData(-1, -1, null)]
        [InlineData(0, 0, "")]
        [InlineData(2, 2, "abc")]
        public void Ctor_ColumnIndex_StringValue(int columnIndex, int rowIndex, string stringValue)
        {
            var result = new ReadCellValueResult(columnIndex, rowIndex, stringValue);
            Assert.Equal(columnIndex, result.ColumnIndex);
            Assert.Equal(rowIndex, result.RowIndex);
            Assert.Equal(stringValue, result.StringValue);
        }
    }
}

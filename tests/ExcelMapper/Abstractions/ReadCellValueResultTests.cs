using ExcelDataReader;
using Xunit;

namespace ExcelMapper.Abstractions.Tests;

public class ReadCellResultTests
{
    [Fact]
    public void Ctor_Default()
    {
        var result = new ReadCellResult();
        Assert.Equal(0, result.ColumnIndex);
        Assert.Null(result.StringValue);
        Assert.Null(result.Reader);
        Assert.False(result.PreserveFormatting);
    }

    [Theory]
    [InlineData(-1, null, true)]
    [InlineData(-1, null, false)]
    [InlineData(0, "", true)]
    [InlineData(0, "", false)]
    [InlineData(2, "abc", true)]
    [InlineData(2, "abc", false)]
    public void Ctor_ColumnIndex_StringValue_Bool(int columnIndex, string? stringValue, bool preserveFormatting)
    {
        var result = new ReadCellResult(columnIndex, stringValue, preserveFormatting);
        Assert.Equal(columnIndex, result.ColumnIndex);
        Assert.Equal(stringValue, result.StringValue);
        Assert.Null(result.Reader);
        Assert.Equal(preserveFormatting, result.PreserveFormatting);
    }
    

    [Theory]
    [InlineData(-1, true)]
    [InlineData(-1, false)]
    [InlineData(0, true)]
    [InlineData(0, false)]
    [InlineData(2, true)]
    [InlineData(2, false)]
    public void Ctor_ColumnIndex_IExcelDataReader_Bool_NullReader(int columnIndex, bool preserveFormatting)
    {
        var result = new ReadCellResult(columnIndex, (IExcelDataReader)null!, preserveFormatting);
        Assert.Equal(columnIndex, result.ColumnIndex);
        Assert.Null(result.StringValue);
        Assert.Null(result.Reader);
        Assert.Equal(preserveFormatting, result.PreserveFormatting);
    }
}

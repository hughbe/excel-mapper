using ExcelMapper.Tests;

namespace ExcelMapper.Readers.Tests;

public class ColumnIndexReaderTests
{
    [Theory]
    [InlineData(0)]
    [InlineData(10)]
    public void Ctor_ColumnIndex(int columnIndex)
    {
        var reader = new ColumnIndexReader(columnIndex);
        Assert.Equal(columnIndex, reader.ColumnIndex);
    }

    [Fact]
    public void Ctor_NegativeColumnIndex_ThrowsArgumentOutOfRangeException()
    {
        Assert.Throws<ArgumentOutOfRangeException>("columnIndex", () => new ColumnIndexReader(-1));
    }

    [Fact]
    public void TryGetValue_Invoke_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var reader = new ColumnIndexReader(0);
        Assert.True(reader.TryGetValue(importer.Reader, false, out var result));
        Assert.Equal(0, result.ColumnIndex);
        Assert.Equal("Value", result.StringValue);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(int.MaxValue)]
    public void TryGetValue_InvokeNoMatch_ReturnsNull(int columnIndex)
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var reader = new ColumnIndexReader(columnIndex);
        Assert.False(reader.TryGetValue(importer.Reader, false, out var result));
        Assert.Equal(0, result.ColumnIndex);
        Assert.Null(result.StringValue);
    }
}

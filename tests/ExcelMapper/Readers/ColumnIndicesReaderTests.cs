namespace ExcelMapper.Readers.Tests;

public class ColumnIndicesReaderTests
{
    [Fact]
    public void Ctor_ColumnIndices()
    {
        var columnIndices = new int[] { 0, 1 };
        var reader = new ColumnIndicesReader(columnIndices);
        Assert.Same(columnIndices, reader.ColumnIndices);
    }

    [Fact]
    public void Ctor_NullColumnIndices_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("columnIndices", () => new ColumnIndicesReader(null!));
    }

    [Fact]
    public void Ctor_EmptyColumnNames_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>("columnIndices", () => new ColumnIndicesReader([]));
    }

    [Fact]
    public void Ctor_NegativeValueInColumnIndices_ThrowsArgumentOutOfRangeException()
    {
        Assert.Throws<ArgumentOutOfRangeException>("columnIndices", () => new ColumnIndicesReader([-1]));
    }
}

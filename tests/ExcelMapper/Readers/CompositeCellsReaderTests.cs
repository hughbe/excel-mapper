using ExcelDataReader;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Readers.Tests;

public class CompositeCellsReaderTests
{
    [Fact]
    public void Ctor_Readers()
    {
        var reader1 = new ColumnIndexReader(0);
        var reader2 = new ColumnIndexReader(1);
        ColumnIndexReader[] readers = [reader1, reader2];
        var reader = new CompositeCellsReader(readers);
        Assert.Same(readers, reader.Readers);
    }

    [Fact]
    public void Ctor_NullReaders_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("readers", () => new CompositeCellsReader(null!));
    }

    [Fact]
    public void Ctor_EmptyReaders_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>("readers", () => new CompositeCellsReader());
    }

    [Fact]
    public void Ctor_NullReaderInReaders_ThrowsArgumentException()
    {
        var reader1 = new ColumnIndexReader(0);
        Assert.Throws<ArgumentException>("readers", () => new CompositeCellsReader([reader1, null!]));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void TryGetValues_Invoke_ReturnsExpectedResult(bool preserveFormatting)
    {
        var factory = new CompositeCellsReader(
            new MockCellReader
            {
                TryGetValueAction = (reader, preserveFormatting) => (true, new ReadCellResult(0, "Value1", preserveFormatting))
            },
            new MockCellReader
            {
                TryGetValueAction = (reader, preserveFormatting) => (false, new ReadCellResult())
            },
            new MockCellReader
            {
                TryGetValueAction = (reader, preserveFormatting) => (true, new ReadCellResult(1, "Value2", preserveFormatting))
            }
        );

        for (var i = 0; i < 2; i++)
        {
            Assert.True(factory.Start(null!, preserveFormatting, out var count));
            Assert.Equal(2, count);
            var resultList = new List<ReadCellResult>();
            while (factory.TryGetNext(out var result))
            {
                resultList.Add(result);
            }
            Assert.Equal(2, resultList.Count);
            Assert.Equal(0, resultList[0].ColumnIndex);
            Assert.Equal("Value1", resultList[0].StringValue);
            Assert.Equal(preserveFormatting, resultList[0].PreserveFormatting);
            Assert.Equal(1, resultList[1].ColumnIndex);
            Assert.Equal("Value2", resultList[1].StringValue);
            Assert.Equal(preserveFormatting, resultList[1].PreserveFormatting);

            // Reset for the next iteration.
            factory.Reset();
        }
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void TryGetValues_InvokeNoResults_ReturnsExpectedResult(bool preserveFormatting)
    {
        var factory = new CompositeCellsReader(
            new MockCellReader
            {
                TryGetValueAction = (reader, preserveFormatting) => (false, new ReadCellResult())
            }
        );

        Assert.True(factory.Start(null!, preserveFormatting, out var count));
        Assert.Equal(0, count);
        Assert.False(factory.TryGetNext(out var result));
    }

    private class MockCellReader : ICellReader
    {
        public Func<IExcelDataReader, bool, (bool, ReadCellResult)> TryGetValueAction { get; set; } = default!;

        public bool TryGetValue(IExcelDataReader reader, bool preserveFormatting, out ReadCellResult result)
        {
            var (success, cellResult) = TryGetValueAction(reader, preserveFormatting);
            result = cellResult!;
            return success;
        }
    }
}

using ExcelDataReader;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Readers.Tests;

public class SplitCellValueReaderTests
{
    [Fact]
    public void Ctor_ICellReaderFactory()
    {
        var readerFactory = new ColumnNameReaderFactory("ColumnName");
        var factory = new SubSplitReaderFactory(readerFactory);
        Assert.Same(readerFactory, factory.ReaderFactory);

        Assert.Equal(StringSplitOptions.None, factory.Options);
    }

    [Fact]
    public void Ctor_NullReaderFactory_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("readerFactory", () => new SubSplitReaderFactory(null!));
    }

    [Theory]
    [InlineData(StringSplitOptions.None - 1)]
    [InlineData(StringSplitOptions.None)]
    [InlineData(StringSplitOptions.RemoveEmptyEntries)]
    [InlineData(StringSplitOptions.RemoveEmptyEntries + 1)]
    public void Options_Set_GetReturnsExpected(StringSplitOptions options)
    {
        var factory = new SubSplitReaderFactory(new ColumnNameReaderFactory("ColumnName"))
        {
            Options = options
        };
        Assert.Equal(options, factory.Options);

        // Set same.
        factory.Options = options;
        Assert.Equal(options, factory.Options);
    }

    [Fact]
    public void ReaderFactory_SetValid_GetReturnsExpected()
    {
        var value = new ColumnNameReaderFactory("ColumnName1");
        var factory = new SubSplitReaderFactory(new ColumnNameReaderFactory("ColumnName2"))
        {
            ReaderFactory = value
        };

        Assert.Same(value, factory.ReaderFactory);

        // Set same.
        factory.ReaderFactory = value;
        Assert.Same(value, factory.ReaderFactory);
    }

    [Fact]
    public void ReaderFactory_SetNull_ThrowsArgumentNullException()
    {
        var factory = new SubSplitReaderFactory(new ColumnNameReaderFactory("ColumnName"));
        Assert.Throws<ArgumentNullException>("value", () => factory.ReaderFactory = null!);
    }

    [Fact]
    public void GetReader_Invoke_ReturnsExpected()
    {
        var factory = new SubSplitReaderFactory(new MockReaderFactory(new MockReader(() => (true, new ReadCellResult(0, string.Empty, preserveFormatting: false)))));
        var reader = factory.GetCellsReader(null!);
        Assert.NotNull(reader);
        Assert.NotSame(reader, factory.GetCellsReader(null!));
    }

    [Fact]
    public void GetReader_TryGetValuesNullReaderResult_ReturnsEmpty()
    {
        var factory = new SubSplitReaderFactory(new MockReaderFactory(new MockReader(() => (true, new ReadCellResult(0, (string?)null, preserveFormatting: false)))));
        var reader = factory.GetCellsReader(null!)!;
        Assert.True(reader.TryGetValues(null!, false, out IEnumerable<ReadCellResult>? result));
        Assert.Empty(result);
    }

    [Fact]
    public void GetReader_TryGetValuesEmptyReaderResult_ReturnsEmpty()
    {
        var factory = new SubSplitReaderFactory(new MockReaderFactory(new MockReader(() => (true, new ReadCellResult(0, "", preserveFormatting: false)))));
        var reader = factory.GetCellsReader(null!)!;
        Assert.True(reader.TryGetValues(null!, false, out IEnumerable<ReadCellResult>? result));
        Assert.Empty(result);
    }

    [Fact]
    public void GetReader_TryGetValuesNullReader_ReturnsNull()
    {
        var factory = new SubSplitReaderFactory(new MockReaderFactory(null));
        Assert.Null(factory.GetCellsReader(null!));
    }

    [Fact]
    public void GetReader_TryGetValuesFalseReader_ReturnsEmpty()
    {
        var factory = new SubSplitReaderFactory(new MockReaderFactory(new MockReader(() => (false, default))));
        var reader = factory.GetCellsReader(null!)!;
        Assert.False(reader.TryGetValues(null!, false, out IEnumerable<ReadCellResult>? result));
        Assert.Null(result);
    }

    private class SubSplitReaderFactory(ICellReaderFactory readerFactory) : SplitReaderFactory(readerFactory)
    {
        protected override string[] GetValues(string value) => [];
    }

    private class MockReader(Func<(bool, ReadCellResult)> action) : ICellReader
    {
        public bool TryGetValue(IExcelDataReader factory, bool preserveFormatting, out ReadCellResult result)
        {   
            var (ret, res) = action();
            result = res;
            return ret;
        }
    }

    private class MockReaderFactory(ICellReader? result) : ICellReaderFactory
    {
        public ICellReader? GetCellReader(ExcelSheet sheet) => result;
    }
}

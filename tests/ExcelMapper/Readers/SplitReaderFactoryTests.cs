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
    public void GetCellsReader_Invoke_ReturnsExpected()
    {
        var innerReaderFactory = new MockCellReaderFactory
        {
            GetCellReaderAction = sheet => new MockReader(() => (true, new ReadCellResult(0, "value1,value2", preserveFormatting: false)))
        };
        var factory = new SubSplitReaderFactory(innerReaderFactory);
        var reader = factory.GetCellsReader(null!);
        Assert.NotNull(reader);
        Assert.NotSame(reader, factory.GetCellsReader(null!));
    }

    [Fact]
    public void GetCellsReader_TryGetValuesNullReaderResult_ReturnsEmpty()
    {
        var factory = new SubSplitReaderFactory(new MockCellReaderFactory
        {
            GetCellReaderAction = sheet => new MockReader(() => (true, new ReadCellResult(0, (string?)null, preserveFormatting: false)))
        });
        var reader = factory.GetCellsReader(null!)!;
        Assert.True(reader.TryGetValues(null!, false, out IEnumerable<ReadCellResult>? result));
        Assert.Empty(result);
    }

    [Fact]
    public void GetCellsReader_TryGetValuesEmptyReaderResult_ReturnsEmpty()
    {
        var innerReaderFactory = new MockCellReaderFactory
        {
            GetCellReaderAction = sheet => new MockReader(() => (true, new ReadCellResult(0, string.Empty, preserveFormatting: false)))
        };
        var factory = new SubSplitReaderFactory(innerReaderFactory);
        var reader = factory.GetCellsReader(null!)!;
        Assert.True(reader.TryGetValues(null!, false, out IEnumerable<ReadCellResult>? result));
        Assert.Empty(result);
    }

    [Fact]
    public void GetCellsReader_TryGetValuesNullReader_ReturnsNull()
    {
        var innerReaderFactory = new MockCellReaderFactory
        {
            GetCellReaderAction = sheet => null
        };
        var factory = new SubSplitReaderFactory(innerReaderFactory);
        Assert.Null(factory.GetCellsReader(null!));
    }

    [Fact]
    public void GetCellsReader_TryGetValuesFalseReader_ReturnsEmpty()
    {
        var innerReaderFactory = new MockCellReaderFactory
        {
            GetCellReaderAction = sheet => new MockReader(() => (false, default))
        };
        var factory = new SubSplitReaderFactory(innerReaderFactory);
        var reader = factory.GetCellsReader(null!)!;
        Assert.False(reader.TryGetValues(null!, false, out IEnumerable<ReadCellResult>? result));
        Assert.Null(result);
    }

    [Fact]
    public void GetColumnNames_InvokeReaderFactoryImplementsIColumnNameProviderCellReaderFactory_ReturnsExpected()
    {
        var innerReaderFactory = new MockIColumnNameProviderCellReaderFactory
        {
            GetColumnNameAction = sheet => "ColumnName"
        };
        var factory = new SubSplitReaderFactory(innerReaderFactory);
        Assert.Equal(["ColumnName"], factory.GetColumnNames(null!));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public void GetColumnNames_InvokeReaderFactoryImplementsIColumnNameProviderCellReaderFactoryNull_ReturnsExpected(string? columnName)
    {
        var innerReaderFactory = new MockIColumnNameProviderCellReaderFactory
        {
            GetColumnNameAction = sheet => columnName!
        };
        var factory = new SubSplitReaderFactory(innerReaderFactory);
        Assert.Null(factory.GetColumnNames(null!));
    }

    [Fact]
    public void GetColumnNames_InvokeReaderFactoryNone_ReturnsExpected()
    {
        var innerReaderFactory = new MockCellReaderFactory();
        var factory = new SubSplitReaderFactory(innerReaderFactory);
        Assert.Null(factory.GetColumnNames(null!));
    }

    [Fact]
    public void GetColumnIndices_InvokeReaderFactoryImplementsIColumnIndexProviderCellReaderFactory_ReturnsExpected()
    {
        var innerReaderFactory = new MockIColumnIndexProviderCellReaderFactory
        {
            GetColumnIndexAction = sheet => 1
        };
        var factory = new SubSplitReaderFactory(innerReaderFactory);
        Assert.Equal([1], factory.GetColumnIndices(null!));
    }

    [Theory]
    [InlineData(null)]
    [InlineData(-1)]
    public void GetColumnIndices_InvokeReaderFactoryImplementsIColumnIndexProviderCellReaderFactoryNull_ReturnsExpected(int? columnIndex)
    {
        var innerReaderFactory = new MockIColumnIndexProviderCellReaderFactory
        {
            GetColumnIndexAction = sheet => columnIndex
        };
        var factory = new SubSplitReaderFactory(innerReaderFactory);
        Assert.Null(factory.GetColumnIndices(null!));
    }

    [Fact]
    public void GetColumnIndices_InvokeReaderFactoryNone_ReturnsExpected()
    {
        var innerReaderFactory = new MockCellReaderFactory();
        var factory = new SubSplitReaderFactory(innerReaderFactory);
        Assert.Null(factory.GetColumnIndices(null!));
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

    private class MockCellReaderFactory : ICellReaderFactory
    {
        public Func<ExcelSheet, ICellReader?>? GetCellReaderAction { get; set; } = default!;

        public ICellReader? GetCellReader(ExcelSheet sheet) => GetCellReaderAction!.Invoke(sheet);
    }

    private class MockIColumnIndexProviderCellReaderFactory : MockCellReaderFactory, IColumnIndexProviderCellReaderFactory
    {
        public Func<ExcelSheet, int?>? GetColumnIndexAction { get; set; }

        public int? GetColumnIndex(ExcelSheet sheet) => GetColumnIndexAction!.Invoke(sheet);
    }

    private class MockIColumnNameProviderCellReaderFactory : MockCellReaderFactory, IColumnNameProviderCellReaderFactory
    {
        public Func<ExcelSheet, string>? GetColumnNameAction { get; set; }

        public string GetColumnName(ExcelSheet sheet) => GetColumnNameAction!.Invoke(sheet);
    }
}

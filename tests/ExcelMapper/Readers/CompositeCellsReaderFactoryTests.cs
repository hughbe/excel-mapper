using ExcelDataReader;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Readers.Tests;

public class CompositeCellsReaderFactoryTests
{
    [Fact]
    public void Ctor_Factory()
    {
        var factory1 = new ColumnIndexReaderFactory(0);
        var factory2 = new ColumnIndexReaderFactory(1);
        ICellReaderFactory[] factories = [factory1, factory2];
        var factory = new CompositeCellsReaderFactory(factories);
        Assert.Same(factories, factory.Factories);
    }

    [Fact]
    public void Ctor_NullFactories_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("factories", () => new CompositeCellsReaderFactory(null!));
    }

    [Fact]
    public void Ctor_EmptyFactories_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>("factories", () => new CompositeCellsReaderFactory());
    }

    [Fact]
    public void Ctor_NullFactoryInFactories_ThrowsArgumentException()
    {
        var factory1 = new ColumnIndexReaderFactory(0);
        Assert.Throws<ArgumentException>("factories", () => new CompositeCellsReaderFactory([factory1, null!]));
    }

    [Fact]
    public void GetCellReader_Invoke_ReturnsExpectedResult()
    {
        var reader = new MockCellReader
        {
            TryGetValueAction = (reader, preserveFormatting) => (true, new ReadCellResult(0, "Value1", preserveFormatting))
        };
        var factory = new CompositeCellsReaderFactory(
            new MockCellReaderFactory
            {
                GetCellReaderAction = (sheet) => reader
            }
        );

        Assert.Same(reader, factory.GetCellReader(null!));
    }

    [Fact]
    public void GetCellReader_InvokeNoResults_ReturnsExpectedResult()
    {
        var factory = new CompositeCellsReaderFactory(
            new MockCellReaderFactory
            {
                GetCellReaderAction = (sheet) => null
            }
        );

        Assert.Null(factory.GetCellReader(null!));
    }

    [Fact]
    public void GetCellsReader_InvokeMultipleFactories_ReturnsExpectedResult()
    {
        var reader1 = new MockCellReader
        {
            TryGetValueAction = (reader, preserveFormatting) => (true, new ReadCellResult(0, "Value1", preserveFormatting))
        };
        var reader2 = new MockCellReader
        {
            TryGetValueAction = (reader, preserveFormatting) => (true, new ReadCellResult(1, "Value2", preserveFormatting))
        };
        var factory = new CompositeCellsReaderFactory(
            new MockCellReaderFactory
            {
                GetCellReaderAction = (sheet) => reader1
            },
            new MockCellReaderFactory
            {
                GetCellReaderAction = (sheet) => null
            },
            new MockCellReaderFactory
            {
                GetCellReaderAction = (sheet) => reader2
            }
        );

        var reader = Assert.IsType<CompositeCellsReader>(factory.GetCellsReader(null!));
        Assert.Equal([reader1, reader2], reader.Readers);
    }

    [Fact]
    public void GetCellsReader_InvokeNoResults_ReturnsExpectedResult()
    {
        var factory = new CompositeCellsReaderFactory(
            new MockCellReaderFactory
            {
                GetCellReaderAction = (sheet) => null
            }
        );

        Assert.Null(factory.GetCellsReader(null!));
    }

    [Fact]
    public void GetColumnIndex_Matching_ReturnsExpectedResult()
    {
        var factory = new CompositeCellsReaderFactory(
            new MockCellReaderFactory(),
            new MockIColumnIndexProviderCellReaderFactory
            {
                GetColumnIndexAction = (sheet) => 1
            },
            new MockIColumnIndexProviderCellReaderFactory
            {
                GetColumnIndexAction = (sheet) => 2
            }
        );

        Assert.Equal(1, factory.GetColumnIndex(null!));
    }

    [Fact]
    public void GetColumnIndex_NullMatching_ReturnsExpectedResult()
    {
        var factory = new CompositeCellsReaderFactory(
            new MockCellReaderFactory(),
            new MockIColumnIndexProviderCellReaderFactory
            {
                GetColumnIndexAction = (sheet) => null
            },
            new MockIColumnIndexProviderCellReaderFactory
            {
                GetColumnIndexAction = (sheet) => 2
            }
        );

        Assert.Equal(2, factory.GetColumnIndex(null!));
    }

    [Fact]
    public void GetColumnIndex_MinusOneMatching_ReturnsExpectedResult()
    {
        var factory = new CompositeCellsReaderFactory(
            new MockCellReaderFactory(),
            new MockIColumnIndexProviderCellReaderFactory
            {
                GetColumnIndexAction = (sheet) => -1
            },
            new MockIColumnIndexProviderCellReaderFactory
            {
                GetColumnIndexAction = (sheet) => 2
            }
        );

        Assert.Equal(2, factory.GetColumnIndex(null!));
    }

    [Fact]
    public void GetColumnIndex_NoMatching_ReturnsExpectedResult()
    {
        var factory = new CompositeCellsReaderFactory(
            new MockCellReaderFactory(),
            new MockCellReaderFactory()
        );

        Assert.Null(factory.GetColumnIndex(null!));
    }

    [Fact]
    public void GetColumnName_Matching_ReturnsExpectedResult()
    {
        var factory = new CompositeCellsReaderFactory(
            new MockCellReaderFactory(),
            new MockIColumnNameProviderCellReaderFactory
            {
                GetColumnNameAction = (sheet) => "Column1"
            },
            new MockIColumnNameProviderCellReaderFactory
            {
                GetColumnNameAction = (sheet) => "Column2"
            }
        );

        Assert.Equal("Column1", factory.GetColumnName(null!));
    }

    [Fact]
    public void GetColumnName_NullMatching_ReturnsExpectedResult()
    {
        var factory = new CompositeCellsReaderFactory(
            new MockCellReaderFactory(),
            new MockIColumnNameProviderCellReaderFactory
            {
                GetColumnNameAction = (sheet) => null!
            },
            new MockIColumnNameProviderCellReaderFactory
            {
                GetColumnNameAction = (sheet) => "Column2"
            }
        );

        Assert.Equal("Column2", factory.GetColumnName(null!));
    }

    [Fact]
    public void GetColumnName_NoMatching_ReturnsExpectedResult()
    {
        var factory = new CompositeCellsReaderFactory(
            new MockCellReaderFactory(),
            new MockCellReaderFactory()
        );

        Assert.Empty(factory.GetColumnName(null!));
    }

    [Fact]
    public void GetColumnIndices_Matching_ReturnsExpectedResult()
    {
        var factory = new CompositeCellsReaderFactory(
            new MockCellReaderFactory(),
            new MockIColumnIndexProviderCellReaderFactory
            {
                GetColumnIndexAction = (sheet) => 1
            },
            new MockIColumnIndexProviderCellReaderFactory
            {
                GetColumnIndexAction = (sheet) => 2
            }
        );

        Assert.Equal([1, 2], factory.GetColumnIndices(null!));
    }

    [Fact]
    public void GetColumnIndices_NullMatching_ReturnsExpectedResult()
    {
        var factory = new CompositeCellsReaderFactory(
            new MockCellReaderFactory(),
            new MockIColumnIndexProviderCellReaderFactory
            {
                GetColumnIndexAction = (sheet) => null
            },
            new MockIColumnIndexProviderCellReaderFactory
            {
                GetColumnIndexAction = (sheet) => 2
            }
        );

        Assert.Equal([2], factory.GetColumnIndices(null!));
    }

    [Fact]
    public void GetColumnIndices_MinusOneMatching_ReturnsExpectedResult()
    {
        var factory = new CompositeCellsReaderFactory(
            new MockCellReaderFactory(),
            new MockIColumnIndexProviderCellReaderFactory
            {
                GetColumnIndexAction = (sheet) => -1
            },
            new MockIColumnIndexProviderCellReaderFactory
            {
                GetColumnIndexAction = (sheet) => 2
            }
        );

        Assert.Equal([2], factory.GetColumnIndices(null!));
    }

    [Fact]
    public void GetColumnIndices_NoMatching_ReturnsExpectedResult()
    {
        var factory = new CompositeCellsReaderFactory(
            new MockCellReaderFactory(),
            new MockCellReaderFactory()
        );

        Assert.Null(factory.GetColumnIndices(null!));
    }

    [Fact]
    public void GetColumnNames_Matching_ReturnsExpectedResult()
    {
        var factory = new CompositeCellsReaderFactory(
            new MockCellReaderFactory(),
            new MockIColumnNameProviderCellReaderFactory
            {
                GetColumnNameAction = (sheet) => "Column1"
            },
            new MockIColumnNameProviderCellReaderFactory
            {
                GetColumnNameAction = (sheet) => string.Empty
            },
            new MockIColumnNameProviderCellReaderFactory
            {
                GetColumnNameAction = (sheet) => "Column2"
            }
        );

        Assert.Equal(["Column1", "Column2"], factory.GetColumnNames(null!));
    }

    [Fact]
    public void GetColumnNames_NullMatching_ReturnsExpectedResult()
    {
        var factory = new CompositeCellsReaderFactory(
            new MockCellReaderFactory(),
            new MockIColumnNameProviderCellReaderFactory
            {
                GetColumnNameAction = (sheet) => "Column1"
            },
            new MockIColumnNameProviderCellReaderFactory
            {
                GetColumnNameAction = (sheet) => null!
            },
            new MockIColumnNameProviderCellReaderFactory
            {
                GetColumnNameAction = (sheet) => string.Empty
            },
            new MockIColumnNameProviderCellReaderFactory
            {
                GetColumnNameAction = (sheet) => "Column2"
            }
        );

        Assert.Equal(["Column1", "Column2"], factory.GetColumnNames(null!));
    }

    [Fact]
    public void GetColumnNames_NoMatching_ReturnsExpectedResult()
    {
        var factory = new CompositeCellsReaderFactory(
            new MockCellReaderFactory(),
            new MockCellReaderFactory()
        );

        Assert.Null(factory.GetColumnNames(null!));
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

    private class MockCellReaderFactory : ICellReaderFactory
    {
        public Func<ExcelSheet, ICellReader?>? GetCellReaderAction { get; set; }

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

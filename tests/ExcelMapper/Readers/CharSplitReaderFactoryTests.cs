using ExcelMapper.Abstractions;

namespace ExcelMapper.Readers.Tests;

public class CharSplitReaderFactoryTests
{
    [Fact]
    public void Ctor_ICellReaderFactory()
    {
        var innerReader = new ColumnNameReaderFactory("ColumnName");
        var factory = new CharSplitReaderFactory(innerReader);
        Assert.Same(innerReader, factory.ReaderFactory);

        Assert.Equal(StringSplitOptions.None, factory.Options);
        Assert.Equal([','], factory.Separators);
        Assert.Same(factory.Separators, factory.Separators);
    }

    [Theory]
    [InlineData(new char[] { ',' })]
    [InlineData(new char[] { ',', ';' })]
    public void Separators_SetValid_GetReturnsExpected(char[] value)
    {
        var factory = new CharSplitReaderFactory(new ColumnNameReaderFactory("ColumnName"))
        {
            Separators = value
        };
        Assert.Same(value, factory.Separators);

        // Set same.
        factory.Separators = value;
        Assert.Same(value, factory.Separators);
    }

    [Fact]
    public void Separators_SetNull_ThrowsArgumentNullException()
    {
        var factory = new CharSplitReaderFactory(new ColumnNameReaderFactory("ColumnName"));
        Assert.Throws<ArgumentNullException>("value", () => factory.Separators = null!);
    }

    [Fact]
    public void Separators_SetEmpty_ThrowsArgumentException()
    {
        var factory = new CharSplitReaderFactory(new ColumnNameReaderFactory("ColumnName"));
        Assert.Throws<ArgumentException>("value", () => factory.Separators = []);
    }

    [Theory]
    [InlineData(StringSplitOptions.None - 1)]
    [InlineData(StringSplitOptions.None)]
    [InlineData(StringSplitOptions.RemoveEmptyEntries)]
    [InlineData(StringSplitOptions.RemoveEmptyEntries + 1)]
    public void Options_Set_GetReturnsExpected(StringSplitOptions value)
    {
        var factory = new CharSplitReaderFactory(new ColumnNameReaderFactory("ColumnName"))
        {
            Options = value
        };
        Assert.Equal(value, factory.Options);

        // Set same.
        factory.Options = value;
        Assert.Equal(value, factory.Options);
    }

#pragma warning disable CS0184 // The is operator is being used to test interface implementation
    [Fact]
    public void Interfaces_IColumnNameProviderCellReaderFactory_DoesNotImplement()
    {
        var factory = new CharSplitReaderFactory(new ColumnNameReaderFactory("ColumnName"));
        Assert.False(factory is IColumnNameProviderCellReaderFactory);
    }

    [Fact]
    public void Interfaces_IColumnIndexProviderCellReaderFactory_DoesNotImplement()
    {
        var factory = new CharSplitReaderFactory(new ColumnNameReaderFactory("ColumnName"));
        Assert.False(factory is IColumnIndexProviderCellReaderFactory);
    }

    [Fact]
    public void Interfaces_IColumnNamesProviderCellReaderFactory_DoesImplement()
    {
        var factory = new CharSplitReaderFactory(new ColumnNameReaderFactory("ColumnName"));
        Assert.True(factory is IColumnNamesProviderCellReaderFactory);
    }

    [Fact]
    public void Interfaces_IColumnIndicesProviderCellReaderFactory_DoesImplement()
    {
        var factory = new CharSplitReaderFactory(new ColumnNameReaderFactory("ColumnName"));
        Assert.True(factory is IColumnIndicesProviderCellReaderFactory);
    }
#pragma warning restore CS0184

    [Fact]
    public void GetColumnNames_Invoke_ReturnsExpected()
    {
        var factory = new CharSplitReaderFactory(new ColumnNameReaderFactory("ColumnName"));
        Assert.Equal(["ColumnName"], factory.GetColumnNames(null!)!);
    }

    [Fact]
    public void GetColumnNames_InvokeNotColumnNameProviderReaderFactory_ReturnsNull()
    {
        var factory = new CharSplitReaderFactory(new CustomCellReaderFactory());
        Assert.Null(factory.GetColumnNames(null!));
    }

    [Fact]
    public void GetColumnIndices_Invoke_ReturnsExpected()
    {
        var factory = new CharSplitReaderFactory(new ColumnIndexReaderFactory(5));
        Assert.Equal([5], factory.GetColumnIndices(null!)!);
    }

    [Fact]
    public void GetColumnIndices_InvokeNotColumnIndexProviderReaderFactory_ReturnsNull()
    {
        var factory = new CharSplitReaderFactory(new CustomCellReaderFactory());
        Assert.Null(factory.GetColumnIndices(null!));
    }

    private class CustomCellReaderFactory : ICellReaderFactory
    {
        public ICellReader GetCellReader(ExcelSheet sheet) => throw new NotImplementedException();
    }
}

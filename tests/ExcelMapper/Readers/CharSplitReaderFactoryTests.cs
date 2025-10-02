using System;
using Xunit;

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
    }
}

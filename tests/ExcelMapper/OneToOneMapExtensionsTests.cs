using System;
using ExcelMapper.Readers;
using Xunit;

namespace ExcelMapper.Tests;

public class OneToOneMapExtensionsTests : ExcelClassMap<Helpers.TestClass>
{
    [Fact]
    public void WithColumnName_ValidColumnName_Success()
    {
        OneToOneMap<string> propertyMap = Map(t => t.Value);
        Assert.Same(propertyMap, propertyMap.WithColumnName("ColumnName"));

        var factory = Assert.IsType<ColumnNameReaderFactory>(propertyMap.ReaderFactory);
        Assert.Equal("ColumnName", factory.ColumnName);
    }

    [Fact]
    public void WithColumnNameMatching_ValidColumnName_Success()
    {
        OneToOneMap<string> propertyMap = Map(t => t.Value).WithColumnNameMatching(e => e == "ColumnName");
        Assert.Same(propertyMap, propertyMap.WithColumnNameMatching(e => e == "ColumnName"));

        Assert.IsType<ColumnNameMatchingReaderFactory>(propertyMap.ReaderFactory);
    }

    [Fact]
    public void WithColumnName_OptionalColumn_Success()
    {
        OneToOneMap<string> propertyMap = Map(t => t.Value).MakeOptional();
        Assert.True(propertyMap.Optional);
        Assert.Same(propertyMap, propertyMap.WithColumnName("ColumnName"));
        Assert.True(propertyMap.Optional);

        var innerReader = Assert.IsType<ColumnNameReaderFactory>(propertyMap.ReaderFactory);
        Assert.Equal("ColumnName", innerReader.ColumnName);
    }

    [Fact]
    public void WithColumnName_NullColumnName_ThrowsArgumentNullException()
    {
        OneToOneMap<string> propertyMap = Map(t => t.Value);
        Assert.Throws<ArgumentNullException>("columnName", () => propertyMap.WithColumnName(null!));
    }

    [Fact]
    public void WithColumnName_EmptyColumnName_ThrowsArgumentException()
    {
        OneToOneMap<string> propertyMap = Map(t => t.Value);
        Assert.Throws<ArgumentException>("columnName", () => propertyMap.WithColumnName(string.Empty));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    public void WithColumnIndex_ValidColumnIndex_Success(int columnIndex)
    {
        OneToOneMap<string> propertyMap = Map(t => t.Value);
        Assert.Same(propertyMap, propertyMap.WithColumnIndex(columnIndex));

        var factory = Assert.IsType<ColumnIndexReaderFactory>(propertyMap.ReaderFactory);
        Assert.Equal(columnIndex, factory.ColumnIndex);
    }

    [Fact]
    public void WithColumnIndex_OptionalColumn_Success()
    {
        OneToOneMap<string> propertyMap = Map(t => t.Value).MakeOptional();
        Assert.True(propertyMap.Optional);
        Assert.Same(propertyMap, propertyMap.WithColumnIndex(1));
        Assert.True(propertyMap.Optional);

        var innerReader = Assert.IsType<ColumnIndexReaderFactory>(propertyMap.ReaderFactory);
        Assert.Equal(1, innerReader.ColumnIndex);
    }

    [Fact]
    public void WithColumnIndex_NegativeColumnIndex_ThrowsArgumentOutOfRangeException()
    {
        OneToOneMap<string> propertyMap = Map(t => t.Value);

        Assert.Throws<ArgumentOutOfRangeException>("columnIndex", () => propertyMap.WithColumnIndex(-1));
    }
}

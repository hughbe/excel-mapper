using System;
using System.Collections.Generic;
using ExcelMapper.Abstractions;
using ExcelMapper.Readers;
using Xunit;

namespace ExcelMapper.Tests;

public class OneToOneMapExtensionsTests : ExcelClassMap<Helpers.TestClass>
{
    [Fact]
    public void WithColumnName_ValidColumnName_Success()
    {
        OneToOneMap<string> map = Map(t => t.Value);
        Assert.Same(map, map.WithColumnName("ColumnName"));

        var factory = Assert.IsType<ColumnNameReaderFactory>(map.ReaderFactory);
        Assert.Equal("ColumnName", factory.ColumnName);
    }

    public static IEnumerable<object[]> WithColumnNameMatching_ParamsString_TestData()
    {
        yield return new object[] { new string[] { "ColumnName1" } };
        yield return new object[] { new string[] { "ColumnName1", "ColumnName2" } };
        yield return new object[] { new string[] { " " } };
        yield return new object[] { new string[] { "ColumnName", "ColumnName" } };
    }

    [Theory]
    [MemberData(nameof(WithColumnNameMatching_ParamsString_TestData))]
    public void WithColumnNameMatching_ColumnNames_Success(string[] columnNames)
    {
        OneToOneMap<string> map = Map(t => t.Value);
        Assert.Same(map, map.WithColumnNameMatching(columnNames));
        
        var factory1 = Assert.IsType<ColumnNamesReaderFactory>(map.ReaderFactory);
        Assert.Same(columnNames, factory1.ColumnNames);

        Assert.Same(map, map.WithColumnNameMatching(columnNames));
        var factory2 = Assert.IsType<ColumnNamesReaderFactory>(map.ReaderFactory);
        Assert.Same(columnNames, factory2.ColumnNames);
    }

    [Fact]
    public void WithColumnNameMatching_NullColumnNames_ThrowsArgumentNullException()
    {
        OneToOneMap<string> map = Map(t => t.Value);
        Assert.Throws<ArgumentNullException>("columnNames", () => map.WithColumnNameMatching((string[])null!));
    }

    [Fact]
    public void WithColumnNameMatching_EmptyColumnNames_ThrowsArgumentException()
    {
        OneToOneMap<string> map = Map(t => t.Value);
        Assert.Throws<ArgumentException>("columnNames", () => map.WithColumnNameMatching([]));
    }

    [Fact]
    public void WithColumnNameMatching_NullValueInColumnNames_ThrowsArgumentException()
    {
        OneToOneMap<string> map = Map(t => t.Value);
        Assert.Throws<ArgumentException>("columnNames", () => map.WithColumnNameMatching([null!]));
    }

    [Fact]
    public void WithColumnNameMatching_Predicate_Success()
    {
        Func<string, bool> predicate1 = columnName => columnName == "ColumnName";
        Func<string, bool> predicate2 = columnName => columnName == "ColumnName";
        OneToOneMap<string> map = Map(t => t.Value);
        Assert.Same(map, map.WithColumnNameMatching(predicate1));
        
        var factory1 = Assert.IsType<ColumnsMatchingReaderFactory>(map.ReaderFactory);
        Assert.Same(predicate1, Assert.IsType<PredicateColumnMatcher>(factory1.Matcher).Predicate);

        Assert.Same(map, map.WithColumnNameMatching(predicate2));
        var factory2 = Assert.IsType<ColumnsMatchingReaderFactory>(map.ReaderFactory);
        Assert.Same(predicate2, Assert.IsType<PredicateColumnMatcher>(factory2.Matcher).Predicate);
    }

    [Fact]
    public void WithColumnNameMatching_NullPredicate_ThrowsArgumentNullException()
    {
        OneToOneMap<string> map = Map(t => t.Value);
        Assert.Throws<ArgumentNullException>("predicate", () => map.WithColumnNameMatching((Func<string, bool>)null!));
    }

    [Fact]
    public void WithColumnMatching_IExcelColumnMatcher_Success()
    {
        Func<string, bool> predicate1 = columnName => columnName == "ColumnName";
        Func<string, bool> predicate2 = columnName => columnName == "ColumnName";
        var matcher1 = new PredicateColumnMatcher(predicate1);
        var matcher2 = new PredicateColumnMatcher(predicate1);
        OneToOneMap<string> map = Map(t => t.Value);
        Assert.Same(map, map.WithColumnMatching(matcher1));
        
        var factory1 = Assert.IsType<ColumnsMatchingReaderFactory>(map.ReaderFactory);
        Assert.Same(matcher1, factory1.Matcher);

        Assert.Same(map, map.WithColumnMatching(matcher2));
        var factory2 = Assert.IsType<ColumnsMatchingReaderFactory>(map.ReaderFactory);
        Assert.Same(matcher2, factory2.Matcher);
    }

    [Fact]
    public void WithColumnMatching_NullMatcher_ThrowsArgumentNullException()
    {
        OneToOneMap<string> map = Map(t => t.Value);
        Assert.Throws<ArgumentNullException>("matcher", () => map.WithColumnMatching((IExcelColumnMatcher)null!));
    }

    [Fact]
    public void WithColumnName_OptionalColumn_Success()
    {
        OneToOneMap<string> map = Map(t => t.Value).MakeOptional();
        Assert.True(map.Optional);
        Assert.Same(map, map.WithColumnName("ColumnName"));
        Assert.True(map.Optional);

        var innerReader = Assert.IsType<ColumnNameReaderFactory>(map.ReaderFactory);
        Assert.Equal("ColumnName", innerReader.ColumnName);
    }

    [Fact]
    public void WithColumnName_NullColumnName_ThrowsArgumentNullException()
    {
        OneToOneMap<string> map = Map(t => t.Value);
        Assert.Throws<ArgumentNullException>("columnName", () => map.WithColumnName(null!));
    }

    [Fact]
    public void WithColumnName_EmptyColumnName_ThrowsArgumentException()
    {
        OneToOneMap<string> map = Map(t => t.Value);
        Assert.Throws<ArgumentException>("columnName", () => map.WithColumnName(string.Empty));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    public void WithColumnIndex_ValidColumnIndex_Success(int columnIndex)
    {
        OneToOneMap<string> map = Map(t => t.Value);
        Assert.Same(map, map.WithColumnIndex(columnIndex));

        var factory = Assert.IsType<ColumnIndexReaderFactory>(map.ReaderFactory);
        Assert.Equal(columnIndex, factory.ColumnIndex);
    }

    [Fact]
    public void WithColumnIndex_OptionalColumn_Success()
    {
        OneToOneMap<string> map = Map(t => t.Value).MakeOptional();
        Assert.True(map.Optional);
        Assert.Same(map, map.WithColumnIndex(1));
        Assert.True(map.Optional);

        var innerReader = Assert.IsType<ColumnIndexReaderFactory>(map.ReaderFactory);
        Assert.Equal(1, innerReader.ColumnIndex);
    }

    [Fact]
    public void WithColumnIndex_NegativeColumnIndex_ThrowsArgumentOutOfRangeException()
    {
        OneToOneMap<string> map = Map(t => t.Value);

        Assert.Throws<ArgumentOutOfRangeException>("columnIndex", () => map.WithColumnIndex(-1));
    }
}

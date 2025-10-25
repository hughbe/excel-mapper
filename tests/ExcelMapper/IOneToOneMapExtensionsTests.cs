using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Abstractions;
using ExcelMapper.Readers;

namespace ExcelMapper.Tests;

public class IOneToOneMapExtensionsTests
{
    [Fact]
    public void WithColumnName_String_Success()
    {
        var map = new CustomOneToOneMap();
        Assert.Same(map, map.WithColumnName("ColumnName"));

        var factory = Assert.IsType<ColumnNameReaderFactory>(map.ReaderFactory);
        Assert.Equal("ColumnName", factory.ColumnName);
        Assert.Equal(StringComparison.OrdinalIgnoreCase, factory.Comparison);
    }

    [Theory]
    [InlineData(StringComparison.CurrentCulture)]
    [InlineData(StringComparison.CurrentCultureIgnoreCase)]
    [InlineData(StringComparison.InvariantCulture)]
    [InlineData(StringComparison.InvariantCultureIgnoreCase)]
    [InlineData(StringComparison.Ordinal)]
    [InlineData(StringComparison.OrdinalIgnoreCase)]
    public void WithColumnName_StringStringComparison_Success(StringComparison comparison)
    {
        var map = new CustomOneToOneMap();
        Assert.Same(map, map.WithColumnName("ColumnName", comparison));

        var factory = Assert.IsType<ColumnNameReaderFactory>(map.ReaderFactory);
        Assert.Equal("ColumnName", factory.ColumnName);
        Assert.Equal(comparison, factory.Comparison);
    }

    [Fact]
    public void WithColumnName_NullColumnName_ThrowsArgumentNullException()
    {
        var map = new CustomOneToOneMap();
        Assert.Throws<ArgumentNullException>("columnName", () => map.WithColumnName(null!));
        Assert.Throws<ArgumentNullException>("columnName", () => map.WithColumnName(null!, StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void WithColumnName_EmptyColumnName_ThrowsArgumentException()
    {
        var map = new CustomOneToOneMap();
        Assert.Throws<ArgumentException>("columnName", () => map.WithColumnName(string.Empty));
        Assert.Throws<ArgumentException>("columnName", () => map.WithColumnName(string.Empty, StringComparison.OrdinalIgnoreCase));
    }

    [Theory]
    [InlineData(StringComparison.CurrentCulture - 1)]
    [InlineData(StringComparison.OrdinalIgnoreCase + 1)]
    public void WithColumnName_InvalidStringComparison_ThrowsArgumentOutOfRangeException(StringComparison comparison)
    {
        var map = new CustomOneToOneMap();
        Assert.Throws<ArgumentOutOfRangeException>("comparison", () => map.WithColumnName("ColumnName", comparison));
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
        var map = new CustomOneToOneMap();
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
        var map = new CustomOneToOneMap();
        Assert.Throws<ArgumentNullException>("columnNames", () => map.WithColumnNameMatching((string[])null!));
    }

    [Fact]
    public void WithColumnNameMatching_EmptyColumnNames_ThrowsArgumentException()
    {
        var map = new CustomOneToOneMap();
        Assert.Throws<ArgumentException>("columnNames", () => map.WithColumnNameMatching([]));
    }

    [Fact]
    public void WithColumnNameMatching_NullValueInColumnNames_ThrowsArgumentException()
    {
        var map = new CustomOneToOneMap();
        Assert.Throws<ArgumentException>("columnNames", () => map.WithColumnNameMatching([null!]));
    }

    [Fact]
    public void WithColumnNameMatching_EmptyValueInColumnNames_ThrowsArgumentException()
    {
        var map = new CustomOneToOneMap();
        Assert.Throws<ArgumentException>("columnNames", () => map.WithColumnNameMatching([string.Empty]));
    }

    [Fact]
    public void WithColumnNameMatching_Predicate_Success()
    {
        Func<string, bool> predicate1 = columnName => columnName == "ColumnName";
        Func<string, bool> predicate2 = columnName => columnName == "ColumnName";
        var map = new CustomOneToOneMap();
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
        var map = new CustomOneToOneMap();
        Assert.Throws<ArgumentNullException>("predicate", () => map.WithColumnNameMatching((Func<string, bool>)null!));
    }

    [Fact]
    public void WithColumnMatching_IExcelColumnMatcher_Success()
    {
        Func<string, bool> predicate1 = columnName => columnName == "ColumnName";
        Func<string, bool> predicate2 = columnName => columnName == "ColumnName";
        var matcher1 = new PredicateColumnMatcher(predicate1);
        var matcher2 = new PredicateColumnMatcher(predicate1);
        var map = new CustomOneToOneMap();
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
        var map = new CustomOneToOneMap();
        Assert.Throws<ArgumentNullException>("matcher", () => map.WithColumnMatching((IExcelColumnMatcher)null!));
    }

    [Fact]
    public void WithColumnName_OptionalColumn_Success()
    {
        var map = new CustomOneToOneMap().MakeOptional();
        Assert.True(map.Optional);
        Assert.Same(map, map.WithColumnName("ColumnName"));
        Assert.True(map.Optional);

        var innerReader = Assert.IsType<ColumnNameReaderFactory>(map.ReaderFactory);
        Assert.Equal("ColumnName", innerReader.ColumnName);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    public void WithColumnIndex_ValidColumnIndex_Success(int columnIndex)
    {
        var map = new CustomOneToOneMap();
        Assert.Same(map, map.WithColumnIndex(columnIndex));

        var factory = Assert.IsType<ColumnIndexReaderFactory>(map.ReaderFactory);
        Assert.Equal(columnIndex, factory.ColumnIndex);
    }

    [Fact]
    public void WithColumnIndex_OptionalColumn_Success()
    {
        var map = new CustomOneToOneMap().MakeOptional();
        Assert.True(map.Optional);
        Assert.Same(map, map.WithColumnIndex(1));
        Assert.True(map.Optional);

        var innerReader = Assert.IsType<ColumnIndexReaderFactory>(map.ReaderFactory);
        Assert.Equal(1, innerReader.ColumnIndex);
    }

    [Fact]
    public void WithColumnIndex_NegativeColumnIndex_ThrowsArgumentOutOfRangeException()
    {
        var map = new CustomOneToOneMap();

        Assert.Throws<ArgumentOutOfRangeException>("columnIndex", () => map.WithColumnIndex(-1));
    }
    
    [Fact]
    public void WithReaderFactory_ValidReader_Success()
    {
        var factory = new ColumnNameReaderFactory("ColumnName");
        var map = new CustomOneToOneMap();
        Assert.False(map.Optional);
        Assert.Same(map, map.WithReaderFactory(factory));
        Assert.Same(factory, map.ReaderFactory);
    }

    [Fact]
    public void WithReaderFactory_OptionalColumn_Success()
    {
        var factory = new ColumnNameReaderFactory("ColumnName");
        var map = new CustomOneToOneMap().MakeOptional();
        Assert.True(map.Optional);
        Assert.Same(map, map.WithReaderFactory(factory));
        Assert.True(map.Optional);
        Assert.Same(factory, map.ReaderFactory);
    }

    [Fact]
    public void WithReaderFactory_NullReader_ThrowsArgumentNullException()
    {
        var map = new CustomOneToOneMap();
        Assert.Throws<ArgumentNullException>("readerFactory", () => map.WithReaderFactory(null!));
    }
    
    private class CustomOneToOneMap : IOneToOneMap
    {
        public bool Optional { get; set; }
        public bool PreserveFormatting { get; set; }
        public ICellReaderFactory ReaderFactory { get; set; } = default!;

        public IValuePipeline Pipeline => default!;

        public bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, [NotNullWhen(true)] out object? value)
            => throw new NotImplementedException();
    }
}

using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Abstractions;
using ExcelMapper.Factories;
using ExcelMapper.Readers;

namespace ExcelMapper.Tests;

public class IManyToOneMapExtensionsTests
{
    [Fact]
    public void WithColumnNames_ParamsIReadOnlyListString_Success()
    {
        var columnNames = new string[] { "ColumnName1", "ColumnName2" };
        var factory = new ColumnNamesReaderFactory("Column");
        var dictionaryFactory = new DictionaryFactory<string, string>();
        var map = new ManyToOneDictionaryMap<string, string>(factory, dictionaryFactory).WithColumnNames("ColumnNames");
        Assert.Same(map, map.WithColumnNames(columnNames));

        var valueReader = Assert.IsType<ColumnNamesReaderFactory>(map.ReaderFactory);
        Assert.Same(columnNames, valueReader.ColumnNames);
        Assert.Equal(StringComparison.OrdinalIgnoreCase, valueReader.Comparison);
    }

    [Fact]
    public void WithColumnNames_ParamsIEnumerableString_Success()
    {
        var columnNames = new string[] { "ColumnName1", "ColumnName2" };
        var map = new CustomManyToOneMap().WithColumnNames("ColumnNames");
        Assert.Same(map, map.WithColumnNames((IEnumerable<string>)columnNames));

        var valueReader = Assert.IsType<ColumnNamesReaderFactory>(map.ReaderFactory);
        Assert.Equal(columnNames, valueReader.ColumnNames);
        Assert.Equal(StringComparison.OrdinalIgnoreCase, valueReader.Comparison);
    }

    [Theory]
    [InlineData(StringComparison.CurrentCulture)]
    [InlineData(StringComparison.CurrentCultureIgnoreCase)]
    [InlineData(StringComparison.InvariantCulture)]
    [InlineData(StringComparison.InvariantCultureIgnoreCase)]
    [InlineData(StringComparison.Ordinal)]
    [InlineData(StringComparison.OrdinalIgnoreCase)]
    public void WithColumnNames_IReadOnlyListStringStringComparison_Success(StringComparison comparison)
    {
        var columnNames = new string[] { "ColumnName1", "ColumnName2" };
        var factory = new ColumnNamesReaderFactory("Column");
        var dictionaryFactory = new DictionaryFactory<string, string>();
        var map = new ManyToOneDictionaryMap<string, string>(factory, dictionaryFactory).WithColumnNames("ColumnNames");
        Assert.Same(map, map.WithColumnNames(columnNames, comparison));

        var valueReader = Assert.IsType<ColumnNamesReaderFactory>(map.ReaderFactory);
        Assert.Same(columnNames, valueReader.ColumnNames);
        Assert.Equal(comparison, valueReader.Comparison);
    }

    [Theory]
    [InlineData(StringComparison.CurrentCulture)]
    [InlineData(StringComparison.CurrentCultureIgnoreCase)]
    [InlineData(StringComparison.InvariantCulture)]
    [InlineData(StringComparison.InvariantCultureIgnoreCase)]
    [InlineData(StringComparison.Ordinal)]
    [InlineData(StringComparison.OrdinalIgnoreCase)]
    public void WithColumnNames_IEnumerableStringComparison_Success(StringComparison comparison)
    {
        var columnNames = new string[] { "ColumnName1", "ColumnName2" };
        var map = new CustomManyToOneMap().WithColumnNames("ColumnNames");
        Assert.Same(map, map.WithColumnNames((IEnumerable<string>)columnNames, comparison));

        var valueReader = Assert.IsType<ColumnNamesReaderFactory>(map.ReaderFactory);
        Assert.Equal(columnNames, valueReader.ColumnNames);
        Assert.Equal(comparison, valueReader.Comparison);
    }

    [Fact]
    public void WithColumnNames_NullColumnNames_ThrowsArgumentNullException()
    {
        var factory = new ColumnNamesReaderFactory("Column");
        var dictionaryFactory = new DictionaryFactory<string, string>();
        var map = new ManyToOneDictionaryMap<string, string>(factory, dictionaryFactory).WithColumnNames("ColumnNames");

        Assert.Throws<ArgumentNullException>("columnNames", () => map.WithColumnNames(null!));
        Assert.Throws<ArgumentNullException>("columnNames", () => map.WithColumnNames(null!, StringComparison.OrdinalIgnoreCase));
        Assert.Throws<ArgumentNullException>("columnNames", () => map.WithColumnNames((IEnumerable<string>)null!));
        Assert.Throws<ArgumentNullException>("columnNames", () => map.WithColumnNames((IEnumerable<string>)null!, StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void WithColumnNames_EmptyColumnNames_ThrowsArgumentException()
    {
        var factory = new ColumnNamesReaderFactory("Column");
        var dictionaryFactory = new DictionaryFactory<string, string>();
        var map = new ManyToOneDictionaryMap<string, string>(factory, dictionaryFactory).WithColumnNames("ColumnNames");

        Assert.Throws<ArgumentException>("columnNames", () => map.WithColumnNames([]));
        Assert.Throws<ArgumentException>("columnNames", () => map.WithColumnNames([], StringComparison.OrdinalIgnoreCase));
        Assert.Throws<ArgumentException>("columnNames", () => map.WithColumnNames((IEnumerable<string>)new List<string>()));
        Assert.Throws<ArgumentException>("columnNames", () => map.WithColumnNames((IEnumerable<string>)new List<string>(), StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void WithColumnNames_NullValueInColumnNames_ThrowsArgumentException()
    {
        var factory = new ColumnNamesReaderFactory("Column");
        var dictionaryFactory = new DictionaryFactory<string, string>();
        var map = new ManyToOneDictionaryMap<string, string>(factory, dictionaryFactory).WithColumnNames("ColumnNames");

        Assert.Throws<ArgumentException>("columnNames", () => map.WithColumnNames([null!]));
        Assert.Throws<ArgumentException>("columnNames", () => map.WithColumnNames([null!], StringComparison.OrdinalIgnoreCase));
        Assert.Throws<ArgumentException>("columnNames", () => map.WithColumnNames((IEnumerable<string>)new List<string> { null! }));
        Assert.Throws<ArgumentException>("columnNames", () => map.WithColumnNames((IEnumerable<string>)new List<string> { null! }, StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void WithColumnNames_EmptyValueInColumnNames_ThrowsArgumentException()
    {
        var factory = new ColumnNamesReaderFactory("Column");
        var dictionaryFactory = new DictionaryFactory<string, string>();
        var map = new ManyToOneDictionaryMap<string, string>(factory, dictionaryFactory).WithColumnNames("ColumnNames");

        Assert.Throws<ArgumentException>("columnNames", () => map.WithColumnNames([""]));
        Assert.Throws<ArgumentException>("columnNames", () => map.WithColumnNames([""], StringComparison.OrdinalIgnoreCase));
        Assert.Throws<ArgumentException>("columnNames", () => map.WithColumnNames((IEnumerable<string>)new List<string> { "" }));
        Assert.Throws<ArgumentException>("columnNames", () => map.WithColumnNames((IEnumerable<string>)new List<string> { "" }, StringComparison.OrdinalIgnoreCase));
    }

    [Theory]
    [InlineData(StringComparison.CurrentCulture - 1)]
    [InlineData(StringComparison.OrdinalIgnoreCase + 1)]
    public void WithColumnNames_InvalidStringComparison_ThrowsArgumentOutOfRangeException(StringComparison comparison)
    {
        var columnNames = new string[] { "ColumnName1", "ColumnName2" };
        var factory = new ColumnNamesReaderFactory("Column");
        var dictionaryFactory = new DictionaryFactory<string, string>();
        var map = new ManyToOneDictionaryMap<string, string>(factory, dictionaryFactory);
        Assert.Throws<ArgumentOutOfRangeException>("comparison", () => map.WithColumnNames(columnNames, comparison));
        Assert.Throws<ArgumentOutOfRangeException>("comparison", () => map.WithColumnNames((IEnumerable<string>)columnNames, comparison));
    }

    [Fact]
    public void WithColumnsMatching_Invoke_Success()
    {
        var matcher = new NamesColumnMatcher("ColumnName1", "ColumnName2");
        var map = new CustomManyToOneMap().WithColumnNames("ColumnNames");
        Assert.Same(map, map.WithColumnsMatching(matcher));

        var newFactory = Assert.IsType<ColumnsMatchingReaderFactory>(map.ReaderFactory);
        Assert.Same(matcher, newFactory.Matcher);
    }

    [Fact]
    public void WithColumnsMatching_NullMatcher_ThrowsArgumentNullException()
    {
        var columnIndices = new int[] { 0, 1 };
        var map = new CustomManyToOneMap().WithColumnNames("ColumnNames");

        Assert.Throws<ArgumentNullException>("matcher", () => map.WithColumnsMatching((IExcelColumnMatcher)null!));
    }

    [Fact]
    public void WithColumnIndices_ParamsInt_Success()
    {
        var columnIndices = new int[] { 0, 1 };
        var map = new CustomManyToOneMap().WithColumnNames("ColumnNames");
        Assert.Same(map, map.WithColumnIndices(columnIndices));

        var newFactory = Assert.IsType<ColumnIndicesReaderFactory>(map.ReaderFactory);
        Assert.Same(columnIndices, newFactory.ColumnIndices);
    }

    [Fact]
    public void WithColumnIndices_IEnumerableInt_Success()
    {
        var columnIndices = new List<int> { 0, 1 };
        var map = new CustomManyToOneMap().WithColumnNames("ColumnNames");
        Assert.Same(map, map.WithColumnIndices((IEnumerable<int>)columnIndices));

        var newFactory = Assert.IsType<ColumnIndicesReaderFactory>(map.ReaderFactory);
        Assert.Equal(columnIndices, newFactory.ColumnIndices);
    }

    [Fact]
    public void WithColumnIndices_NullColumnIndices_ThrowsArgumentNullException()
    {
        var map = new CustomManyToOneMap().WithColumnNames("ColumnNames");

        Assert.Throws<ArgumentNullException>("columnIndices", () => map.WithColumnIndices(null!));
        Assert.Throws<ArgumentNullException>("columnIndices", () => map.WithColumnIndices((IEnumerable<int>)null!));
    }

    [Fact]
    public void WithColumnIndices_EmptyColumnIndices_ThrowsArgumentException()
    {
        var map = new CustomManyToOneMap().WithColumnNames("ColumnNames");

        Assert.Throws<ArgumentException>("columnIndices", () => map.WithColumnIndices([]));
        Assert.Throws<ArgumentException>("columnIndices", () => map.WithColumnIndices((IEnumerable<int>)new List<int>()));
    }

    [Fact]
    public void WithColumnIndices_NegativeValueInColumnIndices_ThrowsArgumentOutOfRangeException()
    {
        var map = new CustomManyToOneMap().WithColumnNames("ColumnNames");

        Assert.Throws<ArgumentOutOfRangeException>("columnIndices", () => map.WithColumnIndices([-1]));
        Assert.Throws<ArgumentOutOfRangeException>("columnIndices", () => map.WithColumnIndices((IEnumerable<int>)new List<int> { -1 }));
    }

    [Fact]
    public void WithReaderFactory_OptionalColumn_Success()
    {
        var factory = new ColumnNamesReaderFactory("ColumnName");
        var map = new CustomManyToOneMap().MakeOptional();
        Assert.True(map.Optional);
        Assert.Same(map, map.WithReaderFactory(factory));
        Assert.True(map.Optional);
        Assert.Same(factory, map.ReaderFactory);
    }

    [Fact]
    public void WithColumns_Invoke_SetsReaderFactory()
    {
        var map = new CustomManyToOneMap().MakeOptional();
        Assert.Same(map, map.WithColumns(new FirstColumnSelector()));
        Assert.NotNull(map.ReaderFactory);
        Assert.True(map.Optional);
    }

    [Fact]
    public void WithColumns_NullSelector_ThrowsArgumentNullException()
    {
        var map = new CustomManyToOneMap();
        Assert.Throws<ArgumentNullException>("selector", () => map.WithColumns(null!));
    }

    private sealed class FirstColumnSelector : IColumnSelector
    {
        public IReadOnlyList<int> SelectColumns(ExcelSheet sheet) => [0];

        public string Describe(ExcelSheet sheet) => "the first column";
    }

    [Fact]
    public void WithReaderFactory_NullReader_ThrowsArgumentNullException()
    {
        var map = new CustomManyToOneMap();
        Assert.Throws<ArgumentNullException>("readerFactory", () => map.WithReaderFactory(null!));
    }
    
    private class NamesColumnMatcher : IExcelColumnMatcher
    {
        public string[] ColumnNames { get; }

        public NamesColumnMatcher(params string[] columnNames)
        {
            ColumnNames = columnNames;
        }

        public bool ColumnMatches(ExcelSheet sheet, int columnIndex)
            => ColumnNames.Contains(sheet.Heading!.GetColumnName(columnIndex));
    }
    
    private class CustomManyToOneMap : IManyToOneMap
    {
        public bool Optional { get; set; }
        public bool PreserveFormatting { get; set; }
        public ICellsReaderFactory ReaderFactory { get; set; } = default!;
        public IValuePipeline Pipeline { get; } = default!;

        public bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, [NotNullWhen(true)] out object? value)
            => throw new NotImplementedException();
    }

    [Fact]
    public void WithColumnsMatching_Predicate_Success()
    {
        Func<string, bool> predicate = columnName => columnName.StartsWith("Column");
        var map = new CustomManyToOneMap().WithColumnNames("ColumnNames");
        Assert.Same(map, map.WithColumnsMatching(predicate));

        var factory = Assert.IsType<ColumnsMatchingReaderFactory>(map.ReaderFactory);
        Assert.Same(predicate, Assert.IsType<PredicateColumnMatcher>(factory.Matcher).Predicate);
    }

    [Fact]
    public void WithColumnsMatching_NullPredicate_ThrowsArgumentNullException()
    {
        var map = new CustomManyToOneMap().WithColumnNames("ColumnNames");
        Assert.Throws<ArgumentNullException>("predicate", () => map.WithColumnsMatching((Func<string, bool>)null!));
    }

    [Fact]
    public void WithColumnsMatching_Regex_Success()
    {
        var regex = new System.Text.RegularExpressions.Regex("^Column");
        var map = new CustomManyToOneMap().WithColumnNames("ColumnNames");
        Assert.Same(map, map.WithColumnsMatching(regex));

        var factory = Assert.IsType<ColumnsMatchingReaderFactory>(map.ReaderFactory);
        Assert.Same(regex, Assert.IsType<RegexColumnMatcher>(factory.Matcher).Regex);
    }

    [Fact]
    public void WithColumnsMatching_NullRegex_ThrowsArgumentNullException()
    {
        var map = new CustomManyToOneMap().WithColumnNames("ColumnNames");
        Assert.Throws<ArgumentNullException>("regex", () => map.WithColumnsMatching((System.Text.RegularExpressions.Regex)null!));
    }
}

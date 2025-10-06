using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Abstractions;
using ExcelMapper.Readers;
using Xunit;

namespace ExcelMapper.Tests;

public class IManyToOneMapExtensionsTests
{

    [Fact]
    public void WithColumnNames_ParamsString_Success()
    {
        var columnNames = new string[] { "ColumnName1", "ColumnName2" };
        var factory = new ColumnNamesReaderFactory("Column");
        var valuePipeline = new ValuePipeline<string>();
        CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
        var map = new ManyToOneDictionaryMap<string>(factory, valuePipeline, createDictionaryFactory).WithColumnNames("ColumnNames");
        Assert.Same(map, map.WithColumnNames(columnNames));

        var valueReader = Assert.IsType<ColumnNamesReaderFactory>(map.ReaderFactory);
        Assert.Same(columnNames, valueReader.ColumnNames);
    }

    [Fact]
    public void WithColumnNames_IEnumerableString_Success()
    {
        var columnNames = new string[] { "ColumnName1", "ColumnName2" };
        var map = new CustomManyToOneMap().WithColumnNames("ColumnNames");
        Assert.Same(map, map.WithColumnNames((IEnumerable<string>)columnNames));

        var valueReader = Assert.IsType<ColumnNamesReaderFactory>(map.ReaderFactory);
        Assert.Equal(columnNames, valueReader.ColumnNames);
    }

    [Fact]
    public void WithColumnNames_NullColumnNames_ThrowsArgumentNullException()
    {
        var factory = new ColumnNamesReaderFactory("Column");
        var valuePipeline = new ValuePipeline<string>();
        CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
        var map = new ManyToOneDictionaryMap<string>(factory, valuePipeline, createDictionaryFactory).WithColumnNames("ColumnNames");

        Assert.Throws<ArgumentNullException>("columnNames", () => map.WithColumnNames(null!));
        Assert.Throws<ArgumentNullException>("columnNames", () => map.WithColumnNames((IEnumerable<string>)null!));
    }

    [Fact]
    public void WithColumnNames_EmptyColumnNames_ThrowsArgumentException()
    {
        var factory = new ColumnNamesReaderFactory("Column");
        var valuePipeline = new ValuePipeline<string>();
        CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
        var map = new ManyToOneDictionaryMap<string>(factory, valuePipeline, createDictionaryFactory).WithColumnNames("ColumnNames");

        Assert.Throws<ArgumentException>("columnNames", () => map.WithColumnNames([]));
        Assert.Throws<ArgumentException>("columnNames", () => map.WithColumnNames(new List<string>()));
    }

    [Fact]
    public void WithColumnNames_NullValueInColumnNames_ThrowsArgumentException()
    {
        var factory = new ColumnNamesReaderFactory("Column");
        var valuePipeline = new ValuePipeline<string>();
        CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
        var map = new ManyToOneDictionaryMap<string>(factory, valuePipeline, createDictionaryFactory).WithColumnNames("ColumnNames");

        Assert.Throws<ArgumentException>("columnNames", () => map.WithColumnNames([null!]));
        Assert.Throws<ArgumentException>("columnNames", () => map.WithColumnNames(new List<string> { null! }));
    }

    [Fact]
    public void WithColumnsMatching_Invoke_Success()
    {
        var matcher = new NamesColumnMatcher("ColumnName1", "ColumnName2");
        var columnIndices = new int[] { 0, 1 };
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

        Assert.Throws<ArgumentNullException>("matcher", () => map.WithColumnsMatching(null!));
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
        Assert.Same(map, map.WithColumnIndices(columnIndices));

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
        Assert.Throws<ArgumentException>("columnIndices", () => map.WithColumnIndices(new List<int>()));
    }

    [Fact]
    public void WithColumnIndices_NegativeValueInColumnIndices_ThrowsArgumentOutOfRangeException()
    {
        var map = new CustomManyToOneMap().WithColumnNames("ColumnNames");

        Assert.Throws<ArgumentOutOfRangeException>("columnIndices", () => map.WithColumnIndices([-1]));
        Assert.Throws<ArgumentOutOfRangeException>("columnIndices", () => map.WithColumnIndices(new List<int> { -1 }));
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

        public bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, [NotNullWhen(true)] out object? value)
            => throw new NotImplementedException();
    }
}

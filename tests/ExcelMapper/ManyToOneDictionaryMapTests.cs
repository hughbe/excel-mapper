using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Abstractions;
using ExcelMapper.Factories;
using ExcelMapper.Readers;

namespace ExcelMapper.Tests;

public class ManyToOneDictionaryMapTests
{
    [Fact]
    public void Ctor_MemberInfo_ICellsReader_IPipeline_CreateDictionaryFactory()
    {
        var factory = new ColumnNamesReaderFactory("Column");
        var dictionaryFactory = new DictionaryFactory<string, string>();
        var map = new ManyToOneDictionaryMap<string, string>(factory, dictionaryFactory);
        Assert.False(map.Optional);
        Assert.False(map.PreserveFormatting);
    }

    [Fact]
    public void Ctor_NullCellValuesReader_ThrowsArgumentNullException()
    {
        var dictionaryFactory = new DictionaryFactory<string, string>();
        Assert.Throws<ArgumentNullException>("readerFactory", () => new ManyToOneDictionaryMap<string, string>(null!, dictionaryFactory));
    }

    [Fact]
    public void Ctor_NullDictionaryFactory_ThrowsArgumentNullException()
    {
        var factory = new ColumnNamesReaderFactory("Column");
        Assert.Throws<ArgumentNullException>("dictionaryFactory", () => new ManyToOneDictionaryMap<string, string>(factory, null!));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Optional_Set_GetReturnsExpected(bool value)
    {
        var factory = new ColumnNamesReaderFactory("Column");
        var dictionaryFactory = new DictionaryFactory<string, string>();
        var map = new ManyToOneDictionaryMap<string, string>(factory, dictionaryFactory)
        {
            Optional = value
        };
        Assert.Equal(value, map.Optional);

        // Set same.
        map.Optional = value;
        Assert.Equal(value, map.Optional);

        // Set different.
        map.Optional = !value;
        Assert.Equal(!value, map.Optional);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PreserveFormatting_Set_GetReturnsExpected(bool value)
    {
        var factory = new ColumnNamesReaderFactory("Column");
        var dictionaryFactory = new DictionaryFactory<string, string>();
        var map = new ManyToOneDictionaryMap<string, string>(factory, dictionaryFactory)
        {
            PreserveFormatting = value
        };
        Assert.Equal(value, map.PreserveFormatting);

        // Set same.
        map.PreserveFormatting = value;
        Assert.Equal(value, map.PreserveFormatting);

        // Set different.
        map.PreserveFormatting = !value;
        Assert.Equal(!value, map.PreserveFormatting);
    }

    public static IEnumerable<object[]> CellValuesReader_Set_TestData()
    {
        yield return new object[] { new ColumnNamesReaderFactory("Column") };
    }

    [Theory]
    [MemberData(nameof(CellValuesReader_Set_TestData))]
    public void CellValuesReader_SetValid_GetReturnsExpected(ICellsReaderFactory value)
    {
        var factory = new ColumnNamesReaderFactory("Column");
        var dictionaryFactory = new DictionaryFactory<string, string>();
        var map = new ManyToOneDictionaryMap<string, string>(factory, dictionaryFactory)
        {
            ReaderFactory = value
        };
        Assert.Same(value, map.ReaderFactory);

        // Set same.
        map.ReaderFactory = value;
        Assert.Same(value, map.ReaderFactory);
    }

    [Fact]
    public void CellValuesReader_SetNull_ThrowsArgumentNullException()
    {
        var factory = new ColumnNamesReaderFactory("Column");
        var dictionaryFactory = new DictionaryFactory<string, string>();
        var map = new ManyToOneDictionaryMap<string, string>(factory, dictionaryFactory);
        Assert.Throws<ArgumentNullException>("value", () => map.ReaderFactory = null!);
    }

    [Fact]
    public void WithValueMap_ValidMap_Success()
    {
        var factory = new ColumnNamesReaderFactory("Column");
        var dictionaryFactory = new DictionaryFactory<string, string>();
        var map = new ManyToOneDictionaryMap<string, string>(factory, dictionaryFactory);

        var newPipeline = new ValuePipeline<string>();
        Assert.Same(map, map.WithValueMap(e =>
        {
            Assert.Same(e, map.Pipeline);
            return newPipeline;
        }));
        Assert.Same(newPipeline, map.Pipeline);
    }

    [Fact]
    public void WithValueMap_NullMap_ThrowsArgumentNullException()
    {
        var factory = new ColumnNamesReaderFactory("Column");
        var dictionaryFactory = new DictionaryFactory<string, string>();
        var map = new ManyToOneDictionaryMap<string, string>(factory, dictionaryFactory);

        Assert.Throws<ArgumentNullException>("valueMap", () => map.WithValueMap(null!));
    }

    [Fact]
    public void WithValueMap_MapReturnsNull_ThrowsArgumentNullException()
    {
        var factory = new ColumnNamesReaderFactory("Column");
        var dictionaryFactory = new DictionaryFactory<string, string>();
        var map = new ManyToOneDictionaryMap<string, string>(factory, dictionaryFactory);

        Assert.Throws<ArgumentNullException>("valueMap", () => map.WithValueMap(_ => null!));
    }

    [Fact]
    public void TryGetValue_InvokeCanRead_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var factory = new MockReaderFactory(new MockReader(() => (true, new MockEnumerator([]))));
        var dictionaryFactory = new DictionaryFactory<string, string>();
        var map = new ManyToOneDictionaryMap<string, string>(factory, dictionaryFactory);
        object? result = null;
        Assert.True(map.TryGetValue(sheet, 0, importer.Reader, null, out result));
        Assert.Empty(Assert.IsType<Dictionary<string, string>>(result));
    }

    [Fact]
    public void TryGetValue_InvokeNullSheet_ThrowsArgumentNullException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var factory = new MockReaderFactory(new MockReader(() => (false, null)));
        var dictionaryFactory = new DictionaryFactory<string, string>();
        var map = new ManyToOneDictionaryMap<string, string>(factory, dictionaryFactory);
        MemberInfo member = typeof(TestClass).GetProperty(nameof(TestClass.Value))!;
        object? result = null;
        Assert.Throws<ArgumentNullException>("sheet", () => map.TryGetValue(null!, 0, importer.Reader, member, out result));
        Assert.Null(result);
    }

    [Fact]
    public void TryGetValue_InvokeSheetWithoutHeadingHasHeading_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();

        var factory = new MockReaderFactory(new MockReader(() => (false, null)));
        var dictionaryFactory = new DictionaryFactory<string, string>();
        var map = new ManyToOneDictionaryMap<string, string>(factory, dictionaryFactory);
        MemberInfo member = typeof(TestClass).GetProperty(nameof(TestClass.Value))!;
        object? result = null;
        Assert.Throws<ExcelMappingException>(() => map.TryGetValue(sheet, 0, importer.Reader, member, out result));
        Assert.Null(result);
    }

    [Fact]
    public void TryGetValue_InvokeSheetWithoutHeadingHasNoHeading_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;

        var factory = new MockReaderFactory(new MockReader(() => (false, null)));
        var dictionaryFactory = new DictionaryFactory<string, string>();
        var map = new ManyToOneDictionaryMap<string, string>(factory, dictionaryFactory);
        MemberInfo member = typeof(TestClass).GetProperty(nameof(TestClass.Value))!;
        object? result = null;
        Assert.Throws<ExcelMappingException>(() => map.TryGetValue(sheet, 0, importer.Reader, member, out result));
        Assert.Null(result);
    }

    [Fact]
    public void TryGetValue_InvokeCantReadPropertyInfo_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var factory = new MockReaderFactory(new MockReader(() => (false, null)));
        var dictionaryFactory = new DictionaryFactory<string, string>();
        var map = new ManyToOneDictionaryMap<string, string>(factory, dictionaryFactory);
        MemberInfo member = typeof(TestClass).GetProperty(nameof(TestClass.Value))!;
        object? result = null;
        Assert.Throws<ExcelMappingException>(() => map.TryGetValue(sheet, 0, importer.Reader, member, out result));
        Assert.Null(result);
    }

    [Fact]
    public void TryGetValue_InvokeCantReadFieldInfo_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var factory = new MockReaderFactory(new MockReader(() => (false, null)));
        var dictionaryFactory = new DictionaryFactory<string, string>();
        var map = new ManyToOneDictionaryMap<string, string>(factory, dictionaryFactory);
        MemberInfo member = typeof(TestClass).GetField(nameof(TestClass._field))!;
        object? result = null;
        Assert.Throws<ExcelMappingException>(() => map.TryGetValue(sheet, 0, importer.Reader, member, out result));
        Assert.Null(result);
    }

    [Fact]
    public void TryGetValue_InvokeCantReadEventInfo_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var factory = new MockReaderFactory(new MockReader(() => (false, null)));
        var dictionaryFactory = new DictionaryFactory<string, string>();
        var map = new ManyToOneDictionaryMap<string, string>(factory, dictionaryFactory);
        MemberInfo member = typeof(TestClass).GetEvent(nameof(TestClass.Event))!;
        object? result = null;
        Assert.Throws<ExcelMappingException>(() => map.TryGetValue(sheet, 0, importer.Reader, member, out result));
        Assert.Null(result);
    }

    [Fact]
    public void TryGetValue_InvokeCantReadNullMember_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var factory = new MockReaderFactory(new MockReader(() => (false, null)));
        var dictionaryFactory = new DictionaryFactory<string, string>();
        var map = new ManyToOneDictionaryMap<string, string>(factory, dictionaryFactory);
        object? result = null;
        Assert.Throws<ExcelMappingException>(() => map.TryGetValue(sheet, 0, importer.Reader, null, out result));
        Assert.Null(result);
    }

    [Fact]
    public void TryGetValue_InvokeCantReadNullMemberIColumnNamesProviderCellReaderFactoryEmpty_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var factory = new ColumnsMatchingReaderFactory(new PredicateColumnMatcher(s => false));
        var dictionaryFactory = new DictionaryFactory<string, string>();
        var map = new ManyToOneDictionaryMap<string, string>(factory, dictionaryFactory);
        object? result = null;
        Assert.Throws<ExcelMappingException>(() => map.TryGetValue(sheet, 0, importer.Reader, null, out result));
        Assert.Null(result);
    }

    [Fact]
    public void TryGetValue_InvokeCantReadNullMemberIColumnNamesProviderCellReaderFactorySingle_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var factory = new ColumnNamesReaderFactory("NoSuchColumn");
        var dictionaryFactory = new DictionaryFactory<string, string>();
        var map = new ManyToOneDictionaryMap<string, string>(factory, dictionaryFactory);
        object? result = null;
        Assert.Throws<ExcelMappingException>(() => map.TryGetValue(sheet, 0, importer.Reader, null, out result));
        Assert.Null(result);
    }

    [Fact]
    public void TryGetValue_InvokeCantReadNullMemberIColumnNamesProviderCellReaderFactoryMultiple_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var factory = new ColumnNamesReaderFactory("NoSuchColumn1", "NoSuchColumn2");
        var dictionaryFactory = new DictionaryFactory<string, string>();
        var map = new ManyToOneDictionaryMap<string, string>(factory, dictionaryFactory);
        object? result = null;
        Assert.Throws<ExcelMappingException>(() => map.TryGetValue(sheet, 0, importer.Reader, null, out result));
        Assert.Null(result);
    }

    [Fact]
    public void TryGetValue_InvokeCantReadNullMemberIColumnIndicesProviderCellReaderFactoryEmpty_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var factory = new EmptyColumnIndicesReaderFactory();
        var dictionaryFactory = new DictionaryFactory<string, string>();
        var map = new ManyToOneDictionaryMap<string, string>(factory, dictionaryFactory);
        object? result = null;
        Assert.Throws<ExcelMappingException>(() => map.TryGetValue(sheet, 0, importer.Reader, null, out result));
        Assert.Null(result);
    }

    [Fact]
    public void TryGetValue_InvokeCantReadNullMemberIColumnIndicesProviderCellReaderFactorySingle_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var factory = new ColumnIndicesReaderFactory(10);
        var dictionaryFactory = new DictionaryFactory<string, string>();
        var map = new ManyToOneDictionaryMap<string, string>(factory, dictionaryFactory);
        object? result = null;
        Assert.Throws<ExcelMappingException>(() => map.TryGetValue(sheet, 0, importer.Reader, null, out result));
        Assert.Null(result);
    }

    [Fact]
    public void TryGetValue_InvokeCantReadNullMemberIColumnIndicesProviderCellReaderFactoryMultiple_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var factory = new ColumnIndicesReaderFactory(0, 1);
        var dictionaryFactory = new DictionaryFactory<string, string>();
        var map = new ManyToOneDictionaryMap<string, string>(factory, dictionaryFactory);
        object? result = null;
        Assert.Throws<ExcelMappingException>(() => map.TryGetValue(sheet, 0, importer.Reader, null, out result));
        Assert.Null(result);
    }

    private class EmptyColumnIndicesReaderFactory : ICellsReaderFactory, IColumnIndicesProviderCellReaderFactory
    {
        public ICellsReader? GetCellsReader(ExcelSheet sheet) => new ColumnIndicesReaderFactory(int.MaxValue).GetCellsReader(sheet);

        public IReadOnlyList<int> GetColumnIndices(ExcelSheet sheet) => [];
    }
    
    private class MockReader : ICellsReader
    {
        private IReadCellResultEnumerator? _enumerator;
        private bool _result;

        public MockReader(Func<(bool, IReadCellResultEnumerator?)> action)
        {
            Action = action;
        }

        public Func<(bool, IReadCellResultEnumerator?)> Action { get; }

        public bool Start(IExcelDataReader reader, bool preserveFormatting, out int count)
        {
            var (ret, res) = Action();
            _result = ret;
            _enumerator = res;
            count = _enumerator?.Count ?? 0;
            return ret;
        }

        public bool TryGetNext([NotNullWhen(true)] out ReadCellResult result)
        {
            if (_enumerator != null && _enumerator.MoveNext())
            {
                result = _enumerator.Current;
                return true;
            }

            result = default;
            return false;
        }

        public void Reset()
        {
            _enumerator?.Reset();
        }
    }

    private class MockReaderFactory(ICellsReader Reader) : ICellsReaderFactory
    {
        public ICellsReader? GetCellsReader(ExcelSheet sheet) => Reader;
    }

    private class MockEnumerator : IReadCellResultEnumerator
    {
        private readonly ReadCellResult[] _results;
        private int _index = -1;

        public MockEnumerator(ReadCellResult[] results)
        {
            _results = results;
        }

        public ReadCellResult Current => _results[_index];

        object? System.Collections.IEnumerator.Current => Current;

        public int Count => _results.Length;

        public void Dispose()
        {
        }

        public bool MoveNext()
        {
            _index++;
            return _index < _results.Length;
        }

        public void Reset()
        {
            _index = -1;
        }
    }

    private class TestClass
    {
        public string Value { get; set; } = default!;
#pragma warning disable 0649
        public string _field = default!;
#pragma warning restore 0649

        public event EventHandler Event { add { } remove { } }
    }
}

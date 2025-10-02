using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Abstractions;
using ExcelMapper.Readers;
using Xunit;

namespace ExcelMapper.Tests;

public class ManyToOneDictionaryMapTests
{
    [Fact]
    public void Ctor_MemberInfo_ICellsReader_IValuePipeline_CreateDictionaryFactory()
    {
        var factory = new ColumnNamesReaderFactory("Column");
        var valuePipeline = new ValuePipeline<string>();
        CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
        var propertyMap = new ManyToOneDictionaryMap<string>(factory, valuePipeline, createDictionaryFactory);
        Assert.NotNull(propertyMap.ValuePipeline);
    }

    [Fact]
    public void Ctor_NullCellValuesReader_ThrowsArgumentNullException()
    {
        var valuePipeline = new ValuePipeline<string>();
        CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
        Assert.Throws<ArgumentNullException>("readerFactory", () => new ManyToOneDictionaryMap<string>(null!, valuePipeline, createDictionaryFactory));
    }

    [Fact]
    public void Ctor_NullPipeline_ThrowsArgumentNullException()
    {
        var factory = new ColumnNamesReaderFactory("Column");
        CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
        Assert.Throws<ArgumentNullException>("valuePipeline", () => new ManyToOneDictionaryMap<string>(factory, null!, createDictionaryFactory));
    }

    [Fact]
    public void Ctor_NullCreateDictionaryFactory_ThrowsArgumentNullException()
    {
        var factory = new ColumnNamesReaderFactory("Column");
        var valuePipeline = new ValuePipeline<string>();
        Assert.Throws<ArgumentNullException>("createDictionaryFactory", () => new ManyToOneDictionaryMap<string>(factory, valuePipeline, null!));
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
        var valuePipeline = new ValuePipeline<string>();
        CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
        var propertyMap = new ManyToOneDictionaryMap<string>(factory, valuePipeline, createDictionaryFactory)
        {
            ReaderFactory = value
        };
        Assert.Same(value, propertyMap.ReaderFactory);

        // Set same.
        propertyMap.ReaderFactory = value;
        Assert.Same(value, propertyMap.ReaderFactory);
    }

    [Fact]
    public void CellValuesReader_SetNull_ThrowsArgumentNullException()
    {
        var factory = new ColumnNamesReaderFactory("Column");
        var valuePipeline = new ValuePipeline<string>();
        CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
        var propertyMap = new ManyToOneDictionaryMap<string>(factory, valuePipeline, createDictionaryFactory);
        Assert.Throws<ArgumentNullException>("value", () => propertyMap.ReaderFactory = null!);
    }

    [Fact]
    public void WithValueMap_ValidMap_Success()
    {
        var factory = new ColumnNamesReaderFactory("Column");
        var valuePipeline = new ValuePipeline<string>();
        CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
        var propertyMap = new ManyToOneDictionaryMap<string>(factory, valuePipeline, createDictionaryFactory);

        var newValuePipeline = new ValuePipeline<string>();
        Assert.Same(propertyMap, propertyMap.WithValueMap(e =>
        {
            Assert.Same(e, propertyMap.ValuePipeline);
            return newValuePipeline;
        }));
        Assert.Same(newValuePipeline, propertyMap.ValuePipeline);
    }

    [Fact]
    public void WithValueMap_NullMap_ThrowsArgumentNullException()
    {
        var factory = new ColumnNamesReaderFactory("Column");
        var valuePipeline = new ValuePipeline<string>();
        CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
        var propertyMap = new ManyToOneDictionaryMap<string>(factory, valuePipeline, createDictionaryFactory);

        Assert.Throws<ArgumentNullException>("valueMap", () => propertyMap.WithValueMap(null!));
    }

    [Fact]
    public void WithValueMap_MapReturnsNull_ThrowsArgumentNullException()
    {
        var factory = new ColumnNamesReaderFactory("Column");
        var valuePipeline = new ValuePipeline<string>();
        CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
        var propertyMap = new ManyToOneDictionaryMap<string>(factory, valuePipeline, createDictionaryFactory);

        Assert.Throws<ArgumentNullException>("valueMap", () => propertyMap.WithValueMap(_ => null!));
    }

    [Fact]
    public void WithColumnNames_ParamsString_Success()
    {
        var columnNames = new string[] { "ColumnName1", "ColumnName2" };
        var factory = new ColumnNamesReaderFactory("Column");
        var valuePipeline = new ValuePipeline<string>();
        CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
        var propertyMap = new ManyToOneDictionaryMap<string>(factory, valuePipeline, createDictionaryFactory).WithColumnNames("ColumnNames");
        Assert.Same(propertyMap, propertyMap.WithColumnNames(columnNames));

        var valueReader = Assert.IsType<ColumnNamesReaderFactory>(propertyMap.ReaderFactory);
        Assert.Same(columnNames, valueReader.ColumnNames);
    }

    [Fact]
    public void WithColumnNames_IEnumerableString_Success()
    {
        var columnNames = new List<string> { "ColumnName1", "ColumnName2" };
        var factory = new ColumnNamesReaderFactory("Column");
        var valuePipeline = new ValuePipeline<string>();
        CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
        var propertyMap = new ManyToOneDictionaryMap<string>(factory, valuePipeline, createDictionaryFactory).WithColumnNames("ColumnNames");
        Assert.Same(propertyMap, propertyMap.WithColumnNames((IEnumerable<string>)columnNames));

        var valueReader = Assert.IsType<ColumnNamesReaderFactory>(propertyMap.ReaderFactory);
        Assert.Equal(columnNames, valueReader.ColumnNames);
    }

    [Fact]
    public void WithColumnNames_NullColumnNames_ThrowsArgumentNullException()
    {
        var factory = new ColumnNamesReaderFactory("Column");
        var valuePipeline = new ValuePipeline<string>();
        CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
        var propertyMap = new ManyToOneDictionaryMap<string>(factory, valuePipeline, createDictionaryFactory).WithColumnNames("ColumnNames");

        Assert.Throws<ArgumentNullException>("columnNames", () => propertyMap.WithColumnNames(null!));
        Assert.Throws<ArgumentNullException>("columnNames", () => propertyMap.WithColumnNames((IEnumerable<string>)null!));
    }

    [Fact]
    public void WithColumnNames_EmptyColumnNames_ThrowsArgumentException()
    {
        var factory = new ColumnNamesReaderFactory("Column");
        var valuePipeline = new ValuePipeline<string>();
        CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
        var propertyMap = new ManyToOneDictionaryMap<string>(factory, valuePipeline, createDictionaryFactory).WithColumnNames("ColumnNames");

        Assert.Throws<ArgumentException>("columnNames", () => propertyMap.WithColumnNames([]));
        Assert.Throws<ArgumentException>("columnNames", () => propertyMap.WithColumnNames(new List<string>()));
    }

    [Fact]
    public void WithColumnNames_NullValueInColumnNames_ThrowsArgumentException()
    {
        var factory = new ColumnNamesReaderFactory("Column");
        var valuePipeline = new ValuePipeline<string>();
        CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
        var propertyMap = new ManyToOneDictionaryMap<string>(factory, valuePipeline, createDictionaryFactory).WithColumnNames("ColumnNames");

        Assert.Throws<ArgumentException>("columnNames", () => propertyMap.WithColumnNames([null!]));
        Assert.Throws<ArgumentException>("columnNames", () => propertyMap.WithColumnNames(new List<string> { null! }));
    }

    [Fact]
    public void TryGetValue_InvokeCanRead_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var factory = new MockReaderFactory(new MockReader(() => (true, [])));
        var valuePipeline = new ValuePipeline<string>();
        CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
        var map = new ManyToOneDictionaryMap<string>(factory, valuePipeline, createDictionaryFactory);
        object? result = null;
        Assert.True(map.TryGetValue(sheet, 0, importer.Reader, null, out result));
        Assert.Empty(Assert.IsType<Dictionary<string, string>>(result));
    }

    [Fact]
    public void TryGetValue_InvokeNullSheet_ThrowsArgumentNullException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var factory = new MockReaderFactory(new MockReader(() => (false, null)));
        var valuePipeline = new ValuePipeline<string>();
        CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
        var map = new ManyToOneDictionaryMap<string>(factory, valuePipeline, createDictionaryFactory);
        MemberInfo member = typeof(TestClass).GetProperty(nameof(TestClass.Value))!;
        object? result = null;
        Assert.Throws<ArgumentNullException>("sheet", () => map.TryGetValue(null!, 0, importer.Reader, member, out result));
        Assert.Null(result);
    }

    [Fact]
    public void TryGetValue_InvokeSheetWithoutHeadingHasHeading_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();

        var factory = new MockReaderFactory(new MockReader(() => (false, null)));
        var valuePipeline = new ValuePipeline<string>();
        CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
        var map = new ManyToOneDictionaryMap<string>(factory, valuePipeline, createDictionaryFactory);
        MemberInfo member = typeof(TestClass).GetProperty(nameof(TestClass.Value))!;
        object? result = null;
        Assert.Throws<ExcelMappingException>(() => map.TryGetValue(sheet, 0, importer.Reader, member, out result));
        Assert.Null(result);
    }

    [Fact]
    public void TryGetValue_InvokeSheetWithoutHeadingHasNoHeading_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();
        sheet.HasHeading = false;

        var factory = new MockReaderFactory(new MockReader(() => (false, null)));
        var valuePipeline = new ValuePipeline<string>();
        CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
        var map = new ManyToOneDictionaryMap<string>(factory, valuePipeline, createDictionaryFactory);
        MemberInfo member = typeof(TestClass).GetProperty(nameof(TestClass.Value))!;
        object? result = null;
        Assert.Throws<ExcelMappingException>(() => map.TryGetValue(sheet, 0, importer.Reader, member, out result));
        Assert.Null(result);
    }

    [Fact]
    public void TryGetValue_InvokeCantReadPropertyInfo_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var factory = new MockReaderFactory(new MockReader(() => (false, null)));
        var valuePipeline = new ValuePipeline<string>();
        CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
        var map = new ManyToOneDictionaryMap<string>(factory, valuePipeline, createDictionaryFactory);
        MemberInfo member = typeof(TestClass).GetProperty(nameof(TestClass.Value))!;
        object? result = null;
        Assert.Throws<ExcelMappingException>(() => map.TryGetValue(sheet, 0, importer.Reader, member, out result));
        Assert.Null(result);
    }

    [Fact]
    public void TryGetValue_InvokeCantReadFieldInfo_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var factory = new MockReaderFactory(new MockReader(() => (false, null)));
        var valuePipeline = new ValuePipeline<string>();
        CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
        var map = new ManyToOneDictionaryMap<string>(factory, valuePipeline, createDictionaryFactory);
        MemberInfo member = typeof(TestClass).GetField(nameof(TestClass._field))!;
        object? result = null;
        Assert.Throws<ExcelMappingException>(() => map.TryGetValue(sheet, 0, importer.Reader, member, out result));
        Assert.Null(result);
    }

    [Fact]
    public void TryGetValue_InvokeCantReadEventInfo_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var factory = new MockReaderFactory(new MockReader(() => (false, null)));
        var valuePipeline = new ValuePipeline<string>();
        CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
        var map = new ManyToOneDictionaryMap<string>(factory, valuePipeline, createDictionaryFactory);
        MemberInfo member = typeof(TestClass).GetEvent(nameof(TestClass.Event))!;
        object? result = null;
        Assert.Throws<ExcelMappingException>(() => map.TryGetValue(sheet, 0, importer.Reader, member, out result));
        Assert.Null(result);
    }

    [Fact]
    public void TryGetValue_InvokeCantReadNullMember_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var factory = new MockReaderFactory(new MockReader(() => (false, null)));
        var valuePipeline = new ValuePipeline<string>();
        CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
        var map = new ManyToOneDictionaryMap<string>(factory, valuePipeline, createDictionaryFactory);
        object? result = null;
        Assert.Throws<ExcelMappingException>(() => map.TryGetValue(sheet, 0, importer.Reader, null, out result));
        Assert.Null(result);
    }
    
    private class MockReader : ICellsReader
    {
        public MockReader(Func<(bool, IEnumerable<ReadCellResult>?)> action)
        {
            Action = action;
        }

        public Func<(bool, IEnumerable<ReadCellResult>?)> Action { get; }

        public bool TryGetValues(IExcelDataReader factory, [NotNullWhen(true)] out IEnumerable<ReadCellResult>? result)
        {
            var (ret, res) = Action();
            result = res;
            return ret;
        }
    }

    private class MockReaderFactory(ICellsReader Reader) : ICellsReaderFactory
    {
        public ICellsReader? GetReader(ExcelSheet sheet) => Reader;
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

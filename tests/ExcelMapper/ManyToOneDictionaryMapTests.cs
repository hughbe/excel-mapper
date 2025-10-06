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
        var map = new ManyToOneDictionaryMap<string>(factory, valuePipeline, createDictionaryFactory);
Assert.False(map.Optional);
        Assert.False(map.PreserveFormatting);
        Assert.NotNull(map.ValuePipeline);
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

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Optional_Set_GetReturnsExpected(bool value)
    {
        var factory = new ColumnNamesReaderFactory("Column");
        var valuePipeline = new ValuePipeline<string>();
        CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
        var map = new ManyToOneDictionaryMap<string>(factory, valuePipeline, createDictionaryFactory)
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
        var valuePipeline = new ValuePipeline<string>();
        CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
        var map = new ManyToOneDictionaryMap<string>(factory, valuePipeline, createDictionaryFactory)
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
        var valuePipeline = new ValuePipeline<string>();
        CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
        var map = new ManyToOneDictionaryMap<string>(factory, valuePipeline, createDictionaryFactory)
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
        var valuePipeline = new ValuePipeline<string>();
        CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
        var map = new ManyToOneDictionaryMap<string>(factory, valuePipeline, createDictionaryFactory);
        Assert.Throws<ArgumentNullException>("value", () => map.ReaderFactory = null!);
    }

    [Fact]
    public void WithValueMap_ValidMap_Success()
    {
        var factory = new ColumnNamesReaderFactory("Column");
        var valuePipeline = new ValuePipeline<string>();
        CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
        var map = new ManyToOneDictionaryMap<string>(factory, valuePipeline, createDictionaryFactory);

        var newValuePipeline = new ValuePipeline<string>();
        Assert.Same(map, map.WithValueMap(e =>
        {
            Assert.Same(e, map.ValuePipeline);
            return newValuePipeline;
        }));
        Assert.Same(newValuePipeline, map.ValuePipeline);
    }

    [Fact]
    public void WithValueMap_NullMap_ThrowsArgumentNullException()
    {
        var factory = new ColumnNamesReaderFactory("Column");
        var valuePipeline = new ValuePipeline<string>();
        CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
        var map = new ManyToOneDictionaryMap<string>(factory, valuePipeline, createDictionaryFactory);

        Assert.Throws<ArgumentNullException>("valueMap", () => map.WithValueMap(null!));
    }

    [Fact]
    public void WithValueMap_MapReturnsNull_ThrowsArgumentNullException()
    {
        var factory = new ColumnNamesReaderFactory("Column");
        var valuePipeline = new ValuePipeline<string>();
        CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
        var map = new ManyToOneDictionaryMap<string>(factory, valuePipeline, createDictionaryFactory);

        Assert.Throws<ArgumentNullException>("valueMap", () => map.WithValueMap(_ => null!));
    }

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
        var columnNames = new List<string> { "ColumnName1", "ColumnName2" };
        var factory = new ColumnNamesReaderFactory("Column");
        var valuePipeline = new ValuePipeline<string>();
        CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
        var map = new ManyToOneDictionaryMap<string>(factory, valuePipeline, createDictionaryFactory).WithColumnNames("ColumnNames");
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
    public void MakeOptional_HasMapper_ReturnsExpected()
    {
        var columnNames = new string[] { "ColumnName1", "ColumnName2" };
        var factory = new ColumnNamesReaderFactory("Column");
        var valuePipeline = new ValuePipeline<string>();
        CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
        var map = new ManyToOneDictionaryMap<string>(factory, valuePipeline, createDictionaryFactory);
        Assert.False(map.Optional);
        Assert.Same(map, map.MakeOptional());
        Assert.True(map.Optional);
        Assert.Same(factory, map.ReaderFactory);
    }

    [Fact]
    public void MakeOptional_AlreadyOptional_ReturnsExpected()
    {
        var columnNames = new string[] { "ColumnName1", "ColumnName2" };
        var factory = new ColumnNamesReaderFactory("Column");
        var valuePipeline = new ValuePipeline<string>();
        CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
        var map = new ManyToOneDictionaryMap<string>(factory, valuePipeline, createDictionaryFactory);
        Assert.Same(map, map.MakeOptional());
        Assert.True(map.Optional);
        Assert.Same(map, map.MakeOptional());
        Assert.True(map.Optional);
        Assert.Same(factory, map.ReaderFactory);
    }

    [Fact]
    public void MakePreserveFormatting_HasMapper_ReturnsExpected()
    {
        var columnNames = new string[] { "ColumnName1", "ColumnName2" };
        var factory = new ColumnNamesReaderFactory("Column");
        var valuePipeline = new ValuePipeline<string>();
        CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
        var map = new ManyToOneDictionaryMap<string>(factory, valuePipeline, createDictionaryFactory);
        Assert.False(map.PreserveFormatting);
        Assert.Same(map, map.MakePreserveFormatting());
        Assert.True(map.PreserveFormatting);
        Assert.Same(factory, map.ReaderFactory);
    }

    [Fact]
    public void MakePreserveFormatting_AlreadyPreserveFormatting_ReturnsExpected()
    {
        var columnNames = new string[] { "ColumnName1", "ColumnName2" };
        var factory = new ColumnNamesReaderFactory("Column");
        var valuePipeline = new ValuePipeline<string>();
        CreateDictionaryFactory<string> createDictionaryFactory = _ => new Dictionary<string, string>();
        var map = new ManyToOneDictionaryMap<string>(factory, valuePipeline, createDictionaryFactory);
        Assert.Same(map, map.MakePreserveFormatting());
        Assert.True(map.PreserveFormatting);
        Assert.Same(map, map.MakePreserveFormatting());
        Assert.True(map.PreserveFormatting);
        Assert.Same(factory, map.ReaderFactory);
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

        public bool TryGetValues(IExcelDataReader factory, bool preserveFormatting, [NotNullWhen(true)] out IEnumerable<ReadCellResult>? result)
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

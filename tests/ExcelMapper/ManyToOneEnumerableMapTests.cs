using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Abstractions;
using ExcelMapper.Factories;
using ExcelMapper.Readers;
using Xunit;

namespace ExcelMapper.Tests;

public class ManyToOneEnumerableMapTests
{
    [Fact]
    public void Ctor_ICellsReaderFactory_IValuePipeline_CreateElementsFactory()
    {
        var factory = new ColumnNamesReaderFactory("Column");
        var elementPipeline = new ValuePipeline<string>();
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(factory, elementPipeline, enumerableFactory);
        Assert.False(map.Optional);
        Assert.False(map.PreserveFormatting);
        Assert.NotNull(map.ElementPipeline);
    }

    [Fact]
    public void Ctor_NullReaderFactory_ThrowsArgumentNullException()
    {
        var elementPipeline = new ValuePipeline<string>();
        var enumerableFactory = new ListEnumerableFactory<string>();
        Assert.Throws<ArgumentNullException>("readerFactory", () => new ManyToOneEnumerableMap<string>(null!, elementPipeline, enumerableFactory));
    }

    [Fact]
    public void Ctor_NullPipeline_ThrowsArgumentNullException()
    {
        var factory = new ColumnNamesReaderFactory("Column");
        var enumerableFactory = new ListEnumerableFactory<string>();
        Assert.Throws<ArgumentNullException>("elementPipeline", () => new ManyToOneEnumerableMap<string>(factory, null!, enumerableFactory));
    }

    [Fact]
    public void Ctor_NullCreateEnumerableFactory_ThrowsArgumentNullException()
    {
        var factory = new ColumnNamesReaderFactory("Column");
        var elementPipeline = new ValuePipeline<string>();
        Assert.Throws<ArgumentNullException>("enumerableFactory", () => new ManyToOneEnumerableMap<string>(factory, elementPipeline, null!));
    }

    public static IEnumerable<object[]> ReaderFactory_Set_TestData()
    {
        yield return new object[] { new ColumnNamesReaderFactory("Column") };
    }

    [Theory]
    [MemberData(nameof(ReaderFactory_Set_TestData))]
    public void ReaderFactory_SetValid_GetReturnsExpected(ICellsReaderFactory value)
    {
        var factory = new ColumnNamesReaderFactory("Column");
        var elementPipeline = new ValuePipeline<string>();
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(factory, elementPipeline, enumerableFactory)
        {
            ReaderFactory = value
        };
        Assert.Same(value, map.ReaderFactory);

        // Set same.
        map.ReaderFactory = value;
        Assert.Same(value, map.ReaderFactory);
    }

    [Fact]
    public void ReaderFactory_SetNull_ThrowsArgumentNullException()
    {
        var factory = new ColumnNamesReaderFactory("Column");
        var elementPipeline = new ValuePipeline<string>();
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(factory, elementPipeline, enumerableFactory);
        Assert.Throws<ArgumentNullException>("value", () => map.ReaderFactory = null!);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Optional_Set_GetReturnsExpected(bool value)
    {
        var factory = new ColumnNamesReaderFactory("Column");
        var elementPipeline = new ValuePipeline<string>();
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(factory, elementPipeline, enumerableFactory)
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
        var elementPipeline = new ValuePipeline<string>();
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(factory, elementPipeline, enumerableFactory)
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

    [Fact]
    public void WithElementMap_ValidMap_Success()
    {
        var factory = new ColumnNamesReaderFactory("Column");
        var elementPipeline = new ValuePipeline<string>();
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(factory, elementPipeline, enumerableFactory);

        var newElementPipeline = new ValuePipeline<string>();
        Assert.Same(map, map.WithElementMap(e =>
        {
            Assert.Same(e, map.ElementPipeline);
            return newElementPipeline;
        }));
        Assert.Same(newElementPipeline, map.ElementPipeline);
    }

    [Fact]
    public void WithElementMap_NullMap_ThrowsArgumentNullException()
    {
        var factory = new ColumnNamesReaderFactory("Column");
        var elementPipeline = new ValuePipeline<string>();
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(factory, elementPipeline, enumerableFactory);

        Assert.Throws<ArgumentNullException>("elementMap", () => map.WithElementMap(null!));
    }

    [Fact]
    public void WithElementMap_MapReturnsNull_ThrowsArgumentNullException()
    {
        var factory = new ColumnNamesReaderFactory("Column");
        var elementPipeline = new ValuePipeline<string>();
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(factory, elementPipeline, enumerableFactory);

        Assert.Throws<ArgumentNullException>("elementMap", () => map.WithElementMap(_ => null!));
    }

    [Fact]
    public void WithColumnName_SplitValidColumnName_Success()
    {
        var factory = new CharSplitReaderFactory(new ColumnNameReaderFactory("Column"));
        var elementPipeline = new ValuePipeline<string>();
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(factory, elementPipeline, enumerableFactory);
        Assert.Same(map, map.WithColumnName("ColumnName"));

        var newFactory = Assert.IsType<CharSplitReaderFactory>(map.ReaderFactory);
        ColumnNameReaderFactory innerReader = Assert.IsType<ColumnNameReaderFactory>(newFactory.ReaderFactory);
        Assert.Equal("ColumnName", innerReader.ColumnName);
    }

    [Fact]
    public void WithColumnName_MultiValidColumnName_Success()
    {
        var factory = new ColumnNamesReaderFactory("Column");
        var elementPipeline = new ValuePipeline<string>();
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(factory, elementPipeline, enumerableFactory).WithColumnNames("ColumnName");
        Assert.Same(map, map.WithColumnName("ColumnName"));

        var newFactory = Assert.IsType<CharSplitReaderFactory>(map.ReaderFactory);
        ColumnNameReaderFactory innerReader = Assert.IsType<ColumnNameReaderFactory>(newFactory.ReaderFactory);
        Assert.Equal("ColumnName", innerReader.ColumnName);
    }

    [Fact]
    public void WithColumnName_NullColumnName_ThrowsArgumentNullException()
    {
        var factory = new ColumnNamesReaderFactory("Column");
        var elementPipeline = new ValuePipeline<string>();
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(factory, elementPipeline, enumerableFactory);

        Assert.Throws<ArgumentNullException>("columnName", () => map.WithColumnName(null!));
    }

    [Fact]
    public void WithColumnName_EmptyColumnName_ThrowsArgumentException()
    {
        var factory = new ColumnNamesReaderFactory("Column");
        var elementPipeline = new ValuePipeline<string>();
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(factory, elementPipeline, enumerableFactory);

        Assert.Throws<ArgumentException>("columnName", () => map.WithColumnName(string.Empty));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    public void WithColumnIndex_SplitColumnIndex_Success(int columnIndex)
    {
        var factory = new CharSplitReaderFactory(new ColumnNameReaderFactory("Column"));
        var elementPipeline = new ValuePipeline<string>();
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(factory, elementPipeline, enumerableFactory);
        Assert.Same(map, map.WithColumnIndex(columnIndex));

        var newFactory = Assert.IsType<CharSplitReaderFactory>(map.ReaderFactory);
        var innerReader = Assert.IsType<ColumnIndexReaderFactory>(newFactory.ReaderFactory);
        Assert.Equal(columnIndex, innerReader.ColumnIndex);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    public void WithColumnIndex_MultiColumnIndex_Success(int columnIndex)
    {
        var factory = new ColumnNamesReaderFactory("Column");
        var elementPipeline = new ValuePipeline<string>();
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(factory, elementPipeline, enumerableFactory).WithColumnNames("ColumnName");
        Assert.Same(map, map.WithColumnIndex(columnIndex));

        var newFactory = Assert.IsType<CharSplitReaderFactory>(map.ReaderFactory);
        var innerReader = Assert.IsType<ColumnIndexReaderFactory>(newFactory.ReaderFactory);
        Assert.Equal(columnIndex, innerReader.ColumnIndex);
    }

    [Fact]
    public void WithColumnIndex_NegativeColumnIndex_ThrowsArgumentOutOfRangeException()
    {
        var factory = new ColumnNamesReaderFactory("Column");
        var elementPipeline = new ValuePipeline<string>();
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(factory, elementPipeline, enumerableFactory);

        Assert.Throws<ArgumentOutOfRangeException>("columnIndex", () => map.WithColumnIndex(-1));
    }

    public static IEnumerable<object[]> Separators_Char_TestData()
    {
        yield return new object[] { new char[] { ',' } };
        yield return new object[] { new char[] { ';', '-' } };
        yield return new object[] { new List<char> { ';', '-' } };
    }

    [Theory]
    [MemberData(nameof(Separators_Char_TestData))]
    public void WithSeparators_ParamsChar_Success(IEnumerable<char> separators)
    {
        char[] separatorsArray = [.. separators];

        var factory = new StringSplitReaderFactory(new ColumnNameReaderFactory("Column"));
        var elementPipeline = new ValuePipeline<string>();
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(factory, elementPipeline, enumerableFactory);
        Assert.Same(map, map.WithSeparators(separatorsArray));

        var newFactory = Assert.IsType<CharSplitReaderFactory>(map.ReaderFactory);
        Assert.Same(separatorsArray, newFactory.Separators);
    }

    [Theory]
    [MemberData(nameof(Separators_Char_TestData))]
    public void WithSeparators_IEnumerableChar_Success(ICollection<char> separators)
    {
        var factory = new StringSplitReaderFactory(new ColumnNameReaderFactory("Column"));
        var elementPipeline = new ValuePipeline<string>();
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(factory, elementPipeline, enumerableFactory);
        Assert.Same(map, map.WithSeparators(separators));

        var newFactory = Assert.IsType<CharSplitReaderFactory>(map.ReaderFactory);
        Assert.Equal(separators, newFactory.Separators);
    }

    public static IEnumerable<object[]> Separators_String_TestData()
    {
        yield return new object[] { new string[] { "," } };
        yield return new object[] { new string[] { ";", "-" } };
        yield return new object[] { new List<string> { ";", "-" } };
    }

    [Theory]
    [MemberData(nameof(Separators_String_TestData))]
    public void WithSeparators_ParamsString_Success(IEnumerable<string> separators)
    {
        string[] separatorsArray = [.. separators];

        var factory = new StringSplitReaderFactory(new ColumnNameReaderFactory("Column"));
        var elementPipeline = new ValuePipeline<string>();
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(factory, elementPipeline, enumerableFactory);
        Assert.Same(map, map.WithSeparators(separatorsArray));

        StringSplitReaderFactory newFactory = Assert.IsType<StringSplitReaderFactory>(map.ReaderFactory);
        Assert.Same(separatorsArray, newFactory.Separators);
    }

    [Theory]
    [MemberData(nameof(Separators_String_TestData))]
    public void WithSeparators_IEnumerableString_Success(ICollection<string> separators)
    {
        var factory = new StringSplitReaderFactory(new ColumnNameReaderFactory("Column"));
        var elementPipeline = new ValuePipeline<string>();
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(factory, elementPipeline, enumerableFactory);
        Assert.Same(map, map.WithSeparators(separators));

        StringSplitReaderFactory newFactory = Assert.IsType<StringSplitReaderFactory>(map.ReaderFactory);
        Assert.Equal(separators, newFactory.Separators);
    }

    [Fact]
    public void WithSeparators_NullSeparators_ThrowsArgumentNullException()
    {
        var factory = new StringSplitReaderFactory(new ColumnNameReaderFactory("Column"));
        var elementPipeline = new ValuePipeline<string>();
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(factory, elementPipeline, enumerableFactory);

        Assert.Throws<ArgumentNullException>("separators", () => map.WithSeparators((char[])null!));
        Assert.Throws<ArgumentNullException>("separators", () => map.WithSeparators((IEnumerable<char>)null!));
        Assert.Throws<ArgumentNullException>("separators", () => map.WithSeparators((string[])null!));
        Assert.Throws<ArgumentNullException>("separators", () => map.WithSeparators((IEnumerable<string>)null!));
    }

    [Fact]
    public void WithSeparators_EmptySeparators_ThrowsArgumentException()
    {
        var factory = new StringSplitReaderFactory(new ColumnNameReaderFactory("Column"));
        var elementPipeline = new ValuePipeline<string>();
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(factory, elementPipeline, enumerableFactory);

        Assert.Throws<ArgumentException>("value", () => map.WithSeparators(new char[0]));
        Assert.Throws<ArgumentException>("value", () => map.WithSeparators(new List<char>()));
        Assert.Throws<ArgumentException>("value", () => map.WithSeparators(new string[0]));
        Assert.Throws<ArgumentException>("value", () => map.WithSeparators(new List<string>()));
    }

    [Fact]
    public void WithSeparators_MultiMap_ThrowsExcelMappingException()
    {
        var factory = new ColumnNamesReaderFactory("Column");
        var elementPipeline = new ValuePipeline<string>();
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(factory, elementPipeline, enumerableFactory).WithColumnNames("ColumnNames");

        Assert.Throws<ExcelMappingException>(() => map.WithSeparators(new char[0]));
        Assert.Throws<ExcelMappingException>(() => map.WithSeparators(new List<char>()));
        Assert.Throws<ExcelMappingException>(() => map.WithSeparators(new string[0]));
        Assert.Throws<ExcelMappingException>(() => map.WithSeparators(new List<string>()));
    }

    [Fact]
    public void TryGetValue_InvokeCanRead_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var factory = new MockReaderFactory(new MockReader(() => (true, [])));
        var elementPipeline = new ValuePipeline<string>();
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(factory, elementPipeline, enumerableFactory);
        object? result = null;
        Assert.True(map.TryGetValue(sheet, 0, importer.Reader, null, out result));
        Assert.Empty(Assert.IsType<List<string>>(result));
    }

    [Fact]
    public void TryGetValue_InvokeCanReadMultiple_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var factory = new MockReaderFactory(new MockReader(() => (true, [new ReadCellResult(0, "Value1", false),new ReadCellResult(0, "Value1", false)])));
        var elementPipeline = new ValuePipeline<string>();
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(factory, elementPipeline, enumerableFactory);
        object? result = null;
        Assert.True(map.TryGetValue(sheet, 0, importer.Reader, null, out result));
        Assert.Equal([null, null], Assert.IsType<List<string?>>(result));
    }

    [Fact]
    public void TryGetValue_SheetWithoutHeadingHasHeading_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        var factory = new MockReaderFactory(new MockReader(() => (true, [])));
        var elementPipeline = new ValuePipeline<string>();
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(factory, elementPipeline, enumerableFactory);
        MemberInfo member = typeof(TestClass).GetProperty(nameof(TestClass.Value))!;
        object? result = null;
        Assert.True(map.TryGetValue(sheet, 0, importer.Reader, member, out result));
        Assert.Empty(Assert.IsType<List<string>>(result));
    }

    [Fact]
    public void TryGetValue_SheetWithoutHeadingHasNoHeading_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;

        var factory = new MockReaderFactory(new MockReader(() => (true, [])));
        var elementPipeline = new ValuePipeline<string>();
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(factory, elementPipeline, enumerableFactory);
        MemberInfo member = typeof(TestClass).GetProperty(nameof(TestClass.Value))!;
        object? result = null;
        Assert.True(map.TryGetValue(sheet, 0, importer.Reader, member, out result));
        Assert.Empty(Assert.IsType<List<string>>(result));
    }

    [Fact]
    public void TryGetValue_NullSheet_ThrowsArgumentNullException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var factory = new MockReaderFactory(new MockReader(() => (false, null)));
        var elementPipeline = new ValuePipeline<string>();
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(factory, elementPipeline, enumerableFactory);
        MemberInfo member = typeof(TestClass).GetProperty(nameof(TestClass.Value))!;
        object? result = null;
        Assert.Throws<ArgumentNullException>("sheet", () => map.TryGetValue(null!, 0, importer.Reader, member, out result));
        Assert.Null(result);
    }

    [Fact]
    public void TryGetValue_InvokeCantReadPropertyInfo_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();

        var factory = new MockReaderFactory(new MockReader(() => (false, null)));
        var elementPipeline = new ValuePipeline<string>();
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(factory, elementPipeline, enumerableFactory);
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

        var factory = new MockReaderFactory(new MockReader(() => (false, null)));
        var elementPipeline = new ValuePipeline<string>();
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(factory, elementPipeline, enumerableFactory);
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

        var factory = new MockReaderFactory(new MockReader(() => (false, null)));
        var elementPipeline = new ValuePipeline<string>();
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(factory, elementPipeline, enumerableFactory);
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

        var factory = new MockReaderFactory(new MockReader(() => (false, null)));
        var elementPipeline = new ValuePipeline<string>();
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(factory, elementPipeline, enumerableFactory);
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

        var factory = new ColumnsMatchingReaderFactory(new NamesColumnMatcher("NoSuchColumn"));
        var elementPipeline = new ValuePipeline<string>();
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(factory, elementPipeline, enumerableFactory);
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
        var elementPipeline = new ValuePipeline<string>();
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(factory, elementPipeline, enumerableFactory);
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
        var elementPipeline = new ValuePipeline<string>();
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(factory, elementPipeline, enumerableFactory);
        object? result = null;
        Assert.Throws<ExcelMappingException>(() => map.TryGetValue(sheet, 0, importer.Reader, null, out result));
        Assert.Null(result);
    }

    [Fact]
    public void TryGetValue_InvokeCantReadNullMemberIColumnIndicesProviderCellReaderFactoryEmpty_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        importer.Reader.Read();

        var factory = new EmptyColumnIndicesReaderFactory();
        var elementPipeline = new ValuePipeline<string>();
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(factory, elementPipeline, enumerableFactory);
        object? result = null;
        Assert.Throws<ExcelMappingException>(() => map.TryGetValue(sheet, 0, importer.Reader, null, out result));
        Assert.Null(result);
    }

    [Fact]
    public void TryGetValue_InvokeCantReadNullMemberIColumnIndicesProviderCellReaderFactorySingle_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        importer.Reader.Read();

        var factory = new ColumnIndicesReaderFactory(int.MaxValue);
        var elementPipeline = new ValuePipeline<string>();
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(factory, elementPipeline, enumerableFactory);
        object? result = null;
        Assert.Throws<ExcelMappingException>(() => map.TryGetValue(sheet, 0, importer.Reader, null, out result));
        Assert.Null(result);
    }

    [Fact]
    public void TryGetValue_InvokeCantReadNullMemberIColumnIndicesProviderCellReaderFactoryMultiple_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        importer.Reader.Read();

        var factory = new ColumnIndicesReaderFactory(1, int.MaxValue);
        var elementPipeline = new ValuePipeline<string>();
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(factory, elementPipeline, enumerableFactory);
        object? result = null;
        Assert.Throws<ExcelMappingException>(() => map.TryGetValue(sheet, 0, importer.Reader, null, out result));
        Assert.Null(result);
    }

    private class EmptyColumnIndicesReaderFactory : ICellsReaderFactory, IColumnIndicesProviderCellReaderFactory
    {
        public ICellsReader? GetCellsReader(ExcelSheet sheet) => new ColumnIndicesReaderFactory(int.MaxValue).GetCellsReader(sheet);

        public int[] GetColumnIndices(ExcelSheet sheet) => [];
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
        public ICellsReader? GetCellsReader(ExcelSheet sheet) => Reader;
    }

    private class TestClass
    {
        public string Value { get; set; } = default!;
#pragma warning disable 0649
        public string _field = default!;
#pragma warning restore 0649

        public event EventHandler Event { add { } remove { } }
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
}

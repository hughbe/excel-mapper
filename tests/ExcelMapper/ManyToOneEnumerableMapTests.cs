using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Abstractions;
using ExcelMapper.Factories;
using ExcelMapper.Readers;

namespace ExcelMapper.Tests;

public class ManyToOneEnumerableMapTests
{
    [Fact]
    public void Ctor_ICellsReaderFactory_IValuePipeline_CreateElementsFactory()
    {
        var readerFactory = new ColumnNamesReaderFactory("Column");
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(readerFactory, enumerableFactory);
        Assert.False(map.Optional);
        Assert.False(map.PreserveFormatting);
        Assert.NotNull(map.Pipeline);
    }

    [Fact]
    public void Ctor_NullReaderFactory_ThrowsArgumentNullException()
    {
        var enumerableFactory = new ListEnumerableFactory<string>();
        Assert.Throws<ArgumentNullException>("readerFactory", () => new ManyToOneEnumerableMap<string>(null!, enumerableFactory));
    }

    [Fact]
    public void Ctor_NullCreateEnumerableFactory_ThrowsArgumentNullException()
    {
        var readerFactory = new ColumnNamesReaderFactory("Column");
        Assert.Throws<ArgumentNullException>("enumerableFactory", () => new ManyToOneEnumerableMap<string>(readerFactory, null!));
    }

    public static IEnumerable<object[]> ReaderFactory_Set_TestData()
    {
        yield return new object[] { new ColumnNamesReaderFactory("Column") };
    }

    [Theory]
    [MemberData(nameof(ReaderFactory_Set_TestData))]
    public void ReaderFactory_SetValid_GetReturnsExpected(ICellsReaderFactory value)
    {
        var readerFactory = new ColumnNamesReaderFactory("Column");
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(readerFactory, enumerableFactory)
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
        var readerFactory = new ColumnNamesReaderFactory("Column");
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(readerFactory, enumerableFactory);
        Assert.Throws<ArgumentNullException>("value", () => map.ReaderFactory = null!);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Optional_Set_GetReturnsExpected(bool value)
    {
        var readerFactory = new ColumnNamesReaderFactory("Column");
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(readerFactory, enumerableFactory)
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
        var readerFactory = new ColumnNamesReaderFactory("Column");
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(readerFactory, enumerableFactory)
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
        var readerFactory = new ColumnNamesReaderFactory("Column");
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(readerFactory, enumerableFactory);

        var newPipeline = new ValuePipeline<string>();
        Assert.Same(map, map.WithElementMap(e =>
        {
            Assert.Same(e, map.Pipeline);
            return newPipeline;
        }));
        Assert.Same(newPipeline, map.Pipeline);
    }

    [Fact]
    public void WithElementMap_NullMap_ThrowsArgumentNullException()
    {
        var readerFactory = new ColumnNamesReaderFactory("Column");
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(readerFactory, enumerableFactory);

        Assert.Throws<ArgumentNullException>("elementMap", () => map.WithElementMap(null!));
    }

    [Fact]
    public void WithElementMap_MapReturnsNull_ThrowsArgumentNullException()
    {
        var readerFactory = new ColumnNamesReaderFactory("Column");
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(readerFactory, enumerableFactory);

        Assert.Throws<ArgumentNullException>("elementMap", () => map.WithElementMap(_ => null!));
    }

    [Fact]
    public void WithColumnName_SplitValidColumnName_Success()
    {
        var readerFactory = new CharSplitReaderFactory(new ColumnNameReaderFactory("Column"));
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(readerFactory, enumerableFactory);
        Assert.Same(map, map.WithColumnName("ColumnName"));

        var newFactory = Assert.IsType<CharSplitReaderFactory>(map.ReaderFactory);
        ColumnNameReaderFactory innerReader = Assert.IsType<ColumnNameReaderFactory>(newFactory.ReaderFactory);
        Assert.Equal("ColumnName", innerReader.ColumnName);
    }

    [Fact]
    public void WithColumnName_MultiValidColumnName_Success()
    {
        var readerFactory = new ColumnNamesReaderFactory("Column");
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(readerFactory, enumerableFactory).WithColumnNames("ColumnName");
        Assert.Same(map, map.WithColumnName("ColumnName"));

        var newFactory = Assert.IsType<CharSplitReaderFactory>(map.ReaderFactory);
        ColumnNameReaderFactory innerReader = Assert.IsType<ColumnNameReaderFactory>(newFactory.ReaderFactory);
        Assert.Equal("ColumnName", innerReader.ColumnName);
    }

    [Theory]
    [InlineData(StringComparison.CurrentCulture)]
    [InlineData(StringComparison.CurrentCultureIgnoreCase)]
    [InlineData(StringComparison.InvariantCulture)]
    [InlineData(StringComparison.InvariantCultureIgnoreCase)]
    [InlineData(StringComparison.Ordinal)]
    [InlineData(StringComparison.OrdinalIgnoreCase)]
    public void WithColumnName_SplitValidColumnNameStringComparison_Success(StringComparison comparison)
    {
        var readerFactory = new CharSplitReaderFactory(new ColumnNameReaderFactory("Column"));
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(readerFactory, enumerableFactory);
        Assert.Same(map, map.WithColumnName("ColumnName", comparison));

        var newFactory = Assert.IsType<CharSplitReaderFactory>(map.ReaderFactory);
        ColumnNameReaderFactory innerReader = Assert.IsType<ColumnNameReaderFactory>(newFactory.ReaderFactory);
        Assert.Equal("ColumnName", innerReader.ColumnName);
    }

    [Theory]
    [InlineData(StringComparison.CurrentCulture)]
    [InlineData(StringComparison.CurrentCultureIgnoreCase)]
    [InlineData(StringComparison.InvariantCulture)]
    [InlineData(StringComparison.InvariantCultureIgnoreCase)]
    [InlineData(StringComparison.Ordinal)]
    [InlineData(StringComparison.OrdinalIgnoreCase)]
    public void WithColumnName_MultiValidColumnNameStringComparison_Success(StringComparison comparison)
    {
        var readerFactory = new ColumnNamesReaderFactory("Column");
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(readerFactory, enumerableFactory).WithColumnNames("ColumnName");
        Assert.Same(map, map.WithColumnName("ColumnName", comparison));

        var newFactory = Assert.IsType<CharSplitReaderFactory>(map.ReaderFactory);
        ColumnNameReaderFactory innerReader = Assert.IsType<ColumnNameReaderFactory>(newFactory.ReaderFactory);
        Assert.Equal("ColumnName", innerReader.ColumnName);
    }

    [Fact]
    public void WithColumnName_NullColumnName_ThrowsArgumentNullException()
    {
        var readerFactory = new ColumnNamesReaderFactory("Column");
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(readerFactory, enumerableFactory);

        Assert.Throws<ArgumentNullException>("columnName", () => map.WithColumnName(null!));
    }

    [Fact]
    public void WithColumnName_EmptyColumnName_ThrowsArgumentException()
    {
        var readerFactory = new ColumnNamesReaderFactory("Column");
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(readerFactory, enumerableFactory);

        Assert.Throws<ArgumentException>("columnName", () => map.WithColumnName(string.Empty));
    }

    [Theory]
    [InlineData(StringComparison.CurrentCulture - 1)]
    [InlineData(StringComparison.OrdinalIgnoreCase + 1)]
    public void WithColumnName_InvalidStringComparison_ThrowsArgumentOutOfRangeException(StringComparison comparison)
    {
        var readerFactory = new ColumnNamesReaderFactory("Column");
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(readerFactory, enumerableFactory);
        Assert.Throws<ArgumentOutOfRangeException>("comparison", () => map.WithColumnName("ColumnName", comparison));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    public void WithColumnIndex_SplitColumnIndex_Success(int columnIndex)
    {
        var readerFactory = new CharSplitReaderFactory(new ColumnNameReaderFactory("Column"));
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(readerFactory, enumerableFactory);
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
        var readerFactory = new ColumnNamesReaderFactory("Column");
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(readerFactory, enumerableFactory).WithColumnNames("ColumnName");
        Assert.Same(map, map.WithColumnIndex(columnIndex));

        var newFactory = Assert.IsType<CharSplitReaderFactory>(map.ReaderFactory);
        var innerReader = Assert.IsType<ColumnIndexReaderFactory>(newFactory.ReaderFactory);
        Assert.Equal(columnIndex, innerReader.ColumnIndex);
    }

    [Fact]
    public void WithColumnIndex_NegativeColumnIndex_ThrowsArgumentOutOfRangeException()
    {
        var readerFactory = new ColumnNamesReaderFactory("Column");
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(readerFactory, enumerableFactory);

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

        var readerFactory = new StringSplitReaderFactory(new ColumnNameReaderFactory("Column"));
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(readerFactory, enumerableFactory);
        Assert.Same(map, map.WithSeparators(separatorsArray));

        var newFactory = Assert.IsType<CharSplitReaderFactory>(map.ReaderFactory);
        Assert.Same(separatorsArray, newFactory.Separators);
    }

    [Theory]
    [MemberData(nameof(Separators_Char_TestData))]
    public void WithSeparators_IEnumerableChar_Success(ICollection<char> separators)
    {
        var readerFactory = new StringSplitReaderFactory(new ColumnNameReaderFactory("Column"));
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(readerFactory, enumerableFactory);
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

        var readerFactory = new StringSplitReaderFactory(new ColumnNameReaderFactory("Column"));
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(readerFactory, enumerableFactory);
        Assert.Same(map, map.WithSeparators(separatorsArray));

        StringSplitReaderFactory newFactory = Assert.IsType<StringSplitReaderFactory>(map.ReaderFactory);
        Assert.Same(separatorsArray, newFactory.Separators);
    }

    [Theory]
    [MemberData(nameof(Separators_String_TestData))]
    public void WithSeparators_IEnumerableString_Success(ICollection<string> separators)
    {
        var readerFactory = new StringSplitReaderFactory(new ColumnNameReaderFactory("Column"));
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(readerFactory, enumerableFactory);
        Assert.Same(map, map.WithSeparators(separators));

        StringSplitReaderFactory newFactory = Assert.IsType<StringSplitReaderFactory>(map.ReaderFactory);
        Assert.Equal(separators, newFactory.Separators);
    }

    [Fact]
    public void WithSeparators_NullSeparators_ThrowsArgumentNullException()
    {
        var readerFactory = new StringSplitReaderFactory(new ColumnNameReaderFactory("Column"));
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(readerFactory, enumerableFactory);

        Assert.Throws<ArgumentNullException>("separators", () => map.WithSeparators((char[])null!));
        Assert.Throws<ArgumentNullException>("separators", () => map.WithSeparators((IEnumerable<char>)null!));
        Assert.Throws<ArgumentNullException>("separators", () => map.WithSeparators((string[])null!));
        Assert.Throws<ArgumentNullException>("separators", () => map.WithSeparators((IEnumerable<string>)null!));
    }

    [Fact]
    public void WithSeparators_EmptySeparators_ThrowsArgumentException()
    {
        var readerFactory = new StringSplitReaderFactory(new ColumnNameReaderFactory("Column"));
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(readerFactory, enumerableFactory);

        Assert.Throws<ArgumentException>("separators", () => map.WithSeparators(Array.Empty<char>()));
        Assert.Throws<ArgumentException>("separators", () => map.WithSeparators(new List<char>()));
        Assert.Throws<ArgumentException>("separators", () => map.WithSeparators(Array.Empty<string>()));
        Assert.Throws<ArgumentException>("separators", () => map.WithSeparators(new List<string>()));
    }

    [Fact]
    public void WithSeparators_NullValueInSeparators_ThrowsArgumentException()
    {
        var readerFactory = new StringSplitReaderFactory(new ColumnNameReaderFactory("Column"));
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(readerFactory, enumerableFactory);

        Assert.Throws<ArgumentException>("separators", () => map.WithSeparators([",", null!]));
        Assert.Throws<ArgumentException>("separators", () => map.WithSeparators(new List<string> { ",", null! }));
    }

    [Fact]
    public void WithSeparators_EmptyValueInSeparators_ThrowsArgumentException()
    {
        var readerFactory = new StringSplitReaderFactory(new ColumnNameReaderFactory("Column"));
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(readerFactory, enumerableFactory);

        Assert.Throws<ArgumentException>("separators", () => map.WithSeparators([",", string.Empty]));
        Assert.Throws<ArgumentException>("separators", () => map.WithSeparators(new List<string> { ",", string.Empty }));
    }

    [Fact]
    public void WithSeparators_MultiMap_ThrowsExcelMappingException()
    {
        var readerFactory = new ColumnNamesReaderFactory("Column");
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(readerFactory, enumerableFactory).WithColumnNames("ColumnNames");

        Assert.Throws<ExcelMappingException>(() => map.WithSeparators(['c']));
        Assert.Throws<ExcelMappingException>(() => map.WithSeparators(new List<char> { 'c' }));
        Assert.Throws<ExcelMappingException>(() => map.WithSeparators([","]));
        Assert.Throws<ExcelMappingException>(() => map.WithSeparators(new List<string> { "," }));
    }

    [Fact]
    public void TryGetValue_InvokeCanRead_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var readerFactory = new MockReaderFactory(new MockReader(() => (true, new MockEnumerator([]))));
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(readerFactory, enumerableFactory);
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

        var readerFactory = new MockReaderFactory(new MockReader(() => (true, new MockEnumerator([new ReadCellResult(0, "Value1", false),new ReadCellResult(0, "Value1", false)]))));
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(readerFactory, enumerableFactory);
        object? result = null;
        Assert.True(map.TryGetValue(sheet, 0, importer.Reader, null, out result));
        Assert.Equal([null, null], Assert.IsType<List<string?>>(result));
    }

    [Fact]
    public void TryGetValue_SheetWithoutHeadingHasHeading_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        var readerFactory = new MockReaderFactory(new MockReader(() => (true, new MockEnumerator([]))));
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(readerFactory, enumerableFactory);
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

        var readerFactory = new MockReaderFactory(new MockReader(() => (true, new MockEnumerator([]))));
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(readerFactory, enumerableFactory);
        MemberInfo member = typeof(TestClass).GetProperty(nameof(TestClass.Value))!;
        object? result = null;
        Assert.True(map.TryGetValue(sheet, 0, importer.Reader, member, out result));
        Assert.Empty(Assert.IsType<List<string>>(result));
    }

    [Fact]
    public void TryGetValue_NullSheet_ThrowsArgumentNullException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var readerFactory = new MockReaderFactory(new MockReader(() => (false, null)));
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(readerFactory, enumerableFactory);
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

        var readerFactory = new MockReaderFactory(new MockReader(() => (false, null)));
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(readerFactory, enumerableFactory);
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

        var readerFactory = new MockReaderFactory(new MockReader(() => (false, null)));
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(readerFactory, enumerableFactory);
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

        var readerFactory = new MockReaderFactory(new MockReader(() => (false, null)));
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(readerFactory, enumerableFactory);
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

        var readerFactory = new MockReaderFactory(new MockReader(() => (false, null)));
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(readerFactory, enumerableFactory);
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

        var readerFactory = new ColumnsMatchingReaderFactory(new NamesColumnMatcher("NoSuchColumn"));
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(readerFactory, enumerableFactory);
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

        var readerFactory = new ColumnNamesReaderFactory("NoSuchColumn");
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(readerFactory, enumerableFactory);
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

        var readerFactory = new ColumnNamesReaderFactory("NoSuchColumn1", "NoSuchColumn2");
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(readerFactory, enumerableFactory);
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

        var readerFactory = new EmptyColumnIndicesReaderFactory();
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(readerFactory, enumerableFactory);
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

        var readerFactory = new ColumnIndicesReaderFactory(int.MaxValue);
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(readerFactory, enumerableFactory);
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

        var readerFactory = new ColumnIndicesReaderFactory(1, int.MaxValue);
        var enumerableFactory = new ListEnumerableFactory<string>();
        var map = new ManyToOneEnumerableMap<string>(readerFactory, enumerableFactory);
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

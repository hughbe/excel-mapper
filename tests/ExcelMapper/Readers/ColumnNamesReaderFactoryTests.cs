using ExcelMapper.Tests;

namespace ExcelMapper.Readers.Tests;

public class ColumnNamesReaderFactoryTests
{
    public static IEnumerable<object[]> Ctor_ParamsString_TestData()
    {
        yield return new object[] { new string[] { "ColumnName1" } };
        yield return new object[] { new string[] { "ColumnName1", "ColumnName2" } };
        yield return new object[] { new string[] { " " } };
        yield return new object[] { new string[] { "ColumnName", "ColumnName" } };
    }

    [Theory]
    [MemberData(nameof(Ctor_ParamsString_TestData))]
    public void Ctor_ParamsString(string[] columnNames)
    {
        var reader = new ColumnNamesReaderFactory(columnNames);
        Assert.Equal(StringComparison.OrdinalIgnoreCase, reader.Comparison);
        Assert.Same(columnNames, reader.ColumnNames);
    }

    public static IEnumerable<object[]> Ctor_IReadOnlyListString_StringComparison_TestData()
    {
        yield return new object[] { new string[] { "ColumnName1" }, StringComparison.CurrentCulture };
        yield return new object[] { new string[] { "ColumnName1", "ColumnName2" }, StringComparison.CurrentCultureIgnoreCase };
        yield return new object[] { new string[] { " " }, StringComparison.InvariantCulture };
        yield return new object[] { new string[] { "ColumnName", "ColumnName" }, StringComparison.InvariantCultureIgnoreCase };
        yield return new object[] { new string[] { "ColumnName" }, StringComparison.Ordinal };
        yield return new object[] { new string[] { "ColumnName1", "ColumnName2" }, StringComparison.OrdinalIgnoreCase };
    }

    [Theory]
    [MemberData(nameof(Ctor_IReadOnlyListString_StringComparison_TestData))]
    public void Ctor_IReadOnlyListString_StringComparison(string[] columnNames, StringComparison comparison)
    {
        var reader = new ColumnNamesReaderFactory(columnNames, comparison);
        Assert.Same(columnNames, reader.ColumnNames);
        Assert.Equal(comparison, reader.Comparison);
    }

    [Fact]
    public void Ctor_NullColumnNames_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("columnNames", () => new ColumnNamesReaderFactory(null!));
        Assert.Throws<ArgumentNullException>("columnNames", () => new ColumnNamesReaderFactory(null!, StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Ctor_EmptyColumnNames_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>("columnNames", () => new ColumnNamesReaderFactory([]));
        Assert.Throws<ArgumentException>("columnNames", () => new ColumnNamesReaderFactory([], StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Ctor_NullValueInColumnNames_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>("columnNames", () => new ColumnNamesReaderFactory([null!]));
        Assert.Throws<ArgumentException>("columnNames", () => new ColumnNamesReaderFactory([null!], StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Ctor_EmptyValueInColumnNames_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>("columnNames", () => new ColumnNamesReaderFactory([string.Empty]));
        Assert.Throws<ArgumentException>("columnNames", () => new ColumnNamesReaderFactory([string.Empty], StringComparison.OrdinalIgnoreCase));
    }

    [Theory]
    [InlineData(StringComparison.CurrentCulture - 1)]
    [InlineData(StringComparison.OrdinalIgnoreCase + 1)]
    public void Ctor_InvalidStringComparison_ThrowsArgumentOutOfRangeException(StringComparison comparison)
    {
        Assert.Throws<ArgumentOutOfRangeException>("comparison", () => new ColumnNamesReaderFactory(["ColumnName"], comparison));
    }

    [Fact]
    public void GetCellReader_InvokeColumnNamesSheetWithHeadingMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var factory = new ColumnNamesReaderFactory("NoSuchColumn", "Value");
        var reader = Assert.IsType<ColumnIndexReader>(factory.GetCellReader(sheet));
        Assert.Equal(0, reader.ColumnIndex);
        Assert.NotSame(reader, factory.GetCellReader(sheet));
    }

    [Fact]
    public void GetCellReader_InvokeColumnNamesNoMatch_ReturnsNull()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var factory = new ColumnNamesReaderFactory("NoSuchColumn");
        Assert.Null(factory.GetCellReader(sheet));
    }

    [Fact]
    public void GetCellReader_NullSheet_ThrowsArgumentNullException()
    {
        var factory = new ColumnNamesReaderFactory("Value");
        Assert.Throws<ArgumentNullException>("sheet",   () => factory.GetCellReader(null!));
    }

    [Fact]
    public void GetCellReader_InvokeColumnNamesSheetNoHeadingHasHeading_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();

        var factory = new ColumnNamesReaderFactory("ColumnName");
        Assert.Throws<ExcelMappingException>(() => factory.GetCellReader(sheet));
        Assert.Null(sheet.Heading);
    }

    [Fact]
    public void GetCellReader_InvokeColumnNamesSheetNoHeadingHasNoHeading_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;

        var factory = new ColumnNamesReaderFactory("ColumnName");
        Assert.Throws<ExcelMappingException>(() => factory.GetCellReader(sheet));
        Assert.Null(sheet.Heading);
    }

    public static IEnumerable<object[]> GetCellsReader_TestData()
    {
        yield return new object[] { new string[] { "Value" }, StringComparison.CurrentCulture, new int[] { 0 } };
        yield return new object[] { new string[] { "Value", "Value" }, StringComparison.CurrentCulture, new int[] { 0, 0 } };
        yield return new object[] { new string[] { "Value" }, StringComparison.CurrentCultureIgnoreCase, new int[] { 0 } };
        yield return new object[] { new string[] { "Value", "Value" }, StringComparison.CurrentCultureIgnoreCase, new int[] { 0, 0 } };
        yield return new object[] { new string[] { "vAlUE" }, StringComparison.CurrentCultureIgnoreCase, new int[] { 0 } };
        yield return new object[] { new string[] { "Value" }, StringComparison.InvariantCulture, new int[] { 0 } };
        yield return new object[] { new string[] { "Value", "Value" }, StringComparison.InvariantCulture, new int[] { 0, 0 } };
        yield return new object[] { new string[] { "Value" }, StringComparison.InvariantCultureIgnoreCase, new int[] { 0 } };
        yield return new object[] { new string[] { "Value", "Value" }, StringComparison.InvariantCultureIgnoreCase, new int[] { 0, 0 } };
        yield return new object[] { new string[] { "vAlUE" }, StringComparison.InvariantCultureIgnoreCase, new int[] { 0 } };
        yield return new object[] { new string[] { "Value" }, StringComparison.Ordinal, new int[] { 0 } };
        yield return new object[] { new string[] { "Value", "Value" }, StringComparison.Ordinal, new int[] { 0, 0 } };
        yield return new object[] { new string[] { "Value" }, StringComparison.OrdinalIgnoreCase, new int[] { 0 } };
        yield return new object[] { new string[] { "Value", "Value" }, StringComparison.OrdinalIgnoreCase, new int[] { 0, 0 } };
        yield return new object[] { new string[] { "vAlUE" }, StringComparison.OrdinalIgnoreCase, new int[] { 0 } };
    }

    [Theory]
    [MemberData(nameof(GetCellsReader_TestData))]
    public void GetCellsReader_InvokeSheetWithHeading_ReturnsExpected(string[] columnNames, StringComparison comparison, int[] expectedColumnIndices)
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var factory = new ColumnNamesReaderFactory(columnNames, comparison);
        var reader = Assert.IsType<ColumnIndicesReader>(factory.GetCellsReader(sheet));
        Assert.Equal(expectedColumnIndices, reader.ColumnIndices);
        Assert.NotSame(reader, factory.GetCellsReader(sheet));
    }

    public static IEnumerable<object[]> GetCellsReader_NoSuchColumn_TestData()
    {
        yield return new object[] { new string[] { "NoSuchColumn" }, StringComparison.CurrentCulture };
        yield return new object[] { new string[] { "Value", "NoSuchColumn" }, StringComparison.CurrentCulture };
        yield return new object[] { new string[] { "Value", "VALUE" }, StringComparison.CurrentCulture };
        yield return new object[] { new string[] { "NoSuchColumn" }, StringComparison.CurrentCultureIgnoreCase };
        yield return new object[] { new string[] { "Value", "NoSuchColumn" }, StringComparison.CurrentCultureIgnoreCase };
        yield return new object[] { new string[] { "NoSuchColumn" }, StringComparison.InvariantCulture };
        yield return new object[] { new string[] { "Value", "NoSuchColumn" }, StringComparison.InvariantCulture };
        yield return new object[] { new string[] { "Value", "VALUE" }, StringComparison.InvariantCulture };
        yield return new object[] { new string[] { "NoSuchColumn" }, StringComparison.InvariantCultureIgnoreCase };
        yield return new object[] { new string[] { "Value", "NoSuchColumn" }, StringComparison.InvariantCultureIgnoreCase };
        yield return new object[] { new string[] { "NoSuchColumn" }, StringComparison.Ordinal };
        yield return new object[] { new string[] { "Value", "NoSuchColumn" }, StringComparison.Ordinal };
        yield return new object[] { new string[] { "Value", "VALUE" }, StringComparison.Ordinal };
        yield return new object[] { new string[] { "NoSuchColumn" }, StringComparison.OrdinalIgnoreCase };
        yield return new object[] { new string[] { "Value", "NoSuchColumn" }, StringComparison.OrdinalIgnoreCase };
    }

    [Theory]
    [MemberData(nameof(GetCellsReader_NoSuchColumn_TestData))]
    public void GetCellsReader_InvokeNoMatch_ReturnsNull(string[] columnNames, StringComparison comparison)
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var factory = new ColumnNamesReaderFactory(columnNames, comparison);
        Assert.Null(factory.GetCellsReader(sheet));
    }

    [Fact]
    public void GetCellsReader_NullSheet_ThrowsArgumentNullException()
    {
        var factory = new ColumnNamesReaderFactory("Value");
        Assert.Throws<ArgumentNullException>("sheet", () => factory.GetCellsReader(null!));
    }

    [Fact]
    public void GetCellsReader_InvokeSheetNoHeadingHasHeading_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();

        var factory = new ColumnNamesReaderFactory("Value");
        Assert.Throws<ExcelMappingException>(() => factory.GetCellsReader(sheet));
        Assert.Null(sheet.Heading);
    }

    [Fact]
    public void GetCellsReader_InvokeSheetNoHeadingHasNoHeading_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;

        var factory = new ColumnNamesReaderFactory("Value");
        Assert.Throws<ExcelMappingException>(() => factory.GetCellsReader(sheet));
        Assert.Null(sheet.Heading);
    }

#pragma warning disable CS0184 // The is operator is being used to test interface implementation
    [Fact]
    public void Interfaces_IColumnNameProviderCellReaderFactory_DoesNotImplement()
    {
        var factory = new ColumnNamesReaderFactory("ColumnName");
        Assert.False(factory is IColumnNameProviderCellReaderFactory);
    }

    [Fact]
    public void Interfaces_IColumnIndexProviderCellReaderFactory_DoesNotImplement()
    {
        var factory = new ColumnNamesReaderFactory("ColumnName");
        Assert.False(factory is IColumnIndexProviderCellReaderFactory);
    }

    [Fact]
    public void Interfaces_IColumnNamesProviderCellReaderFactory_Implements()
    {
        var factory = new ColumnNamesReaderFactory("ColumnName");
        Assert.True(factory is IColumnNamesProviderCellReaderFactory);
    }

    [Fact]
    public void Interfaces_IColumnIndicesProviderCellReaderFactory_DoesImplement()
    {
        var factory = new ColumnNamesReaderFactory("ColumnName");
        Assert.False(factory is IColumnIndicesProviderCellReaderFactory);
    }
#pragma warning restore CS0184

    [Fact]
    public void GetColumnNames_Invoke_ReturnsExpected()
    {
        var factory = new ColumnNamesReaderFactory("Column1", "Column2");
        Assert.Equal(["Column1", "Column2"], factory.GetColumnNames(null!));
    }
}

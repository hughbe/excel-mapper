using ExcelMapper.Tests;

namespace ExcelMapper.Readers.Tests;

public class ColumnNameReaderFactoryTests
{
    [Theory]
    [InlineData("ColumnName")]
    [InlineData("  ColumnName  ")]
    [InlineData("  ")]
    public void Ctor_String(string columnName)
    {
        var factory = new ColumnNameReaderFactory(columnName);
        Assert.Equal(columnName, factory.ColumnName);
        Assert.Equal(StringComparison.OrdinalIgnoreCase, factory.Comparison);
    }

    [Theory]
    [InlineData("ColumnName", StringComparison.CurrentCulture)]
    [InlineData("  ColumnName  ", StringComparison.CurrentCultureIgnoreCase)]
    [InlineData("  ", StringComparison.InvariantCulture)]
    [InlineData("ColumnName", StringComparison.InvariantCultureIgnoreCase)]
    [InlineData("ColumnName", StringComparison.Ordinal)]
    [InlineData("ColumnName", StringComparison.OrdinalIgnoreCase)]
    public void Ctor_String_StringComparison(string columnName, StringComparison comparison)
    {
        var factory = new ColumnNameReaderFactory(columnName, comparison);
        Assert.Equal(columnName, factory.ColumnName);
        Assert.Equal(comparison, factory.Comparison);
    }

    [Fact]
    public void Ctor_NullColumnName_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("columnName", () => new ColumnNameReaderFactory(null!));
        Assert.Throws<ArgumentNullException>("columnName", () => new ColumnNameReaderFactory(null!, StringComparison.CurrentCulture));
    }

    [Fact]
    public void Ctor_EmptyColumnName_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>("columnName", () => new ColumnNameReaderFactory(string.Empty));
        Assert.Throws<ArgumentException>("columnName", () => new ColumnNameReaderFactory(string.Empty, StringComparison.CurrentCulture));
    }

    [Theory]
    [InlineData("Value", StringComparison.OrdinalIgnoreCase)]
    [InlineData("vAlUE", StringComparison.OrdinalIgnoreCase)]
    public void GetReader_InvokeSheetWithHeading_ReturnsExpected(string columnName, StringComparison comparison)
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var factory = new ColumnNameReaderFactory(columnName, comparison);
        var reader = Assert.IsType<ColumnIndexReader>(factory.GetCellReader(sheet));
        Assert.Equal(0, reader.ColumnIndex);
        Assert.NotSame(reader, factory.GetCellReader(sheet));
    }

    [Theory]
    [InlineData("VALUE", StringComparison.CurrentCulture)]
    [InlineData("NoSuchColumn", StringComparison.CurrentCultureIgnoreCase)]
    [InlineData("VALUE", StringComparison.InvariantCulture)]
    [InlineData("NoSuchColumn", StringComparison.InvariantCulture)]
    [InlineData("NoSuchColumn", StringComparison.InvariantCultureIgnoreCase)]
    [InlineData("VALUE", StringComparison.Ordinal)]
    [InlineData("NoSuchColumn", StringComparison.Ordinal)]
    [InlineData("NoSuchColumn", StringComparison.OrdinalIgnoreCase)]
    public void GetReader_InvokeNoMatch_ReturnsNull(string columnName, StringComparison comparison)
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var factory = new ColumnNameReaderFactory(columnName, comparison);
        Assert.Null(factory.GetCellReader(sheet));
    }

    [Fact]
    public void GetReader_NullSheet_ThrowsArgumentNullException()
    {
        var factory = new ColumnNameReaderFactory("Value");
        Assert.Throws<ArgumentNullException>("sheet", () => factory.GetCellReader(null!));
    }

    [Fact]
    public void GetReader_InvokeSheetNoHeadingHasHeading_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();

        var factory = new ColumnNameReaderFactory("Value");
        Assert.Throws<ExcelMappingException>(() => factory.GetCellReader(sheet));
        Assert.Null(sheet.Heading);
    }

    [Fact]
    public void GetReader_InvokeSheetNoHeadingHasNoHeading_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;

        var factory = new ColumnNameReaderFactory("Value");
        Assert.Throws<ExcelMappingException>(() => factory.GetCellReader(sheet));
        Assert.Null(sheet.Heading);
    }

    [Fact]
    public void GetColumnName_Invoke_ReturnsExpected()
    {
        var factory = new ColumnNameReaderFactory("ColumnName");
        Assert.Equal("ColumnName", factory.GetColumnName(null!));
    }
}

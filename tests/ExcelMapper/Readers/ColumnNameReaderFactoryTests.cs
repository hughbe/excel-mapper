using System;
using ExcelMapper.Tests;
using Xunit;

namespace ExcelMapper.Readers.Tests;

public class ColumnNameReaderFactoryTests
{
    [Fact]
    public void Ctor_ColumnName()
    {
        var factory = new ColumnNameReaderFactory("ColumnName");
        Assert.Equal("ColumnName", factory.ColumnName);
    }

    [Fact]
    public void Ctor_NullColumnName_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("columnName", () => new ColumnNameReaderFactory(null!));
    }

    [Fact]
    public void Ctor_EmptyColumnName_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>("columnName", () => new ColumnNameReaderFactory(string.Empty));
    }

    [Theory]
    [InlineData("Value")]
    [InlineData("vAlUE")]
    public void GetReader_InvokeSheetWithHeading_ReturnsExpected(string columnName)
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var factory = new ColumnNameReaderFactory(columnName);
        var reader = Assert.IsType<ColumnIndexReader>(factory.GetReader(sheet));
        Assert.Equal(0, reader.ColumnIndex);
        Assert.NotSame(reader, factory.GetReader(sheet));
    }

    [Fact]
    public void GetReader_InvokeNoMatch_ReturnsNull()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var factory = new ColumnNameReaderFactory("NoSuchColumn");
        Assert.Null(factory.GetReader(sheet));
    }

    [Fact]
    public void GetReader_NullSheet_ThrowsArgumentNullException()
    {
        var factory = new ColumnNameReaderFactory("Value");
        Assert.Throws<ArgumentNullException>(() => factory.GetReader(null!));
    }

    [Fact]
    public void GetReader_InvokeSheetNoHeadingHasHeading_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();

        var factory = new ColumnNameReaderFactory("Value");
        Assert.Throws<ExcelMappingException>(() => factory.GetReader(sheet));
        Assert.Null(sheet.Heading);
    }

    [Fact]
    public void GetReader_InvokeSheetNoHeadingHasNoHeading_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();
        sheet.HasHeading = false;

        var factory = new ColumnNameReaderFactory("Value");
        Assert.Throws<ExcelMappingException>(() => factory.GetReader(sheet));
        Assert.Null(sheet.Heading);
    }
}

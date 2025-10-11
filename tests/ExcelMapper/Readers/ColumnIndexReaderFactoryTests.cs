using System;
using ExcelMapper.Tests;
using Xunit;

namespace ExcelMapper.Readers.Tests;

public class ColumnIndexReaderFactoryTests
{
    [Theory]
    [InlineData(0)]
    [InlineData(10)]
    public void Ctor_ColumnIndex(int columnIndex)
    {
        var reader = new ColumnIndexReaderFactory(columnIndex);
        Assert.Equal(columnIndex, reader.ColumnIndex);
    }

    [Fact]
    public void Ctor_NegativeColumnIndex_ThrowsArgumentOutOfRangeException()
    {
        Assert.Throws<ArgumentOutOfRangeException>("columnIndex", () => new ColumnIndexReaderFactory(-1));
    }

    [Theory]
    [InlineData(0)]
    public void GetReader_InvokeSheetWithHeading_ReturnsExpected(int columnIndex)
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var factory = new ColumnIndexReaderFactory(columnIndex);
        var reader = Assert.IsType<ColumnIndexReader>(factory.GetCellReader(sheet));
        Assert.Equal(columnIndex, reader.ColumnIndex);
        Assert.NotSame(reader, factory.GetCellReader(sheet));
    }

    [Theory]
    [InlineData(1)]
    [InlineData(int.MaxValue)]
    public void GetReader_InvokeNoMatch_ReturnsNull(int columnIndex)
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var factory = new ColumnIndexReaderFactory(columnIndex);
        Assert.Null(factory.GetCellReader(sheet));
    }

    [Theory]
    [InlineData(0)]
    public void GetReader_InvokeSheetWithNoHeadingHasHeading_ReturnsExpected(int columnIndex)
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();

        var factory = new ColumnIndexReaderFactory(columnIndex);
        var reader = Assert.IsType<ColumnIndexReader>(factory.GetCellReader(sheet));
        Assert.Equal(columnIndex, reader.ColumnIndex);
        Assert.NotSame(reader, factory.GetCellReader(sheet));
        Assert.Null(sheet.Heading);
    }

    [Theory]
    [InlineData(0)]
    public void GetReader_InvokeSheetWithNoHeadingHasNoHeading_ReturnsExpected(int columnIndex)
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();
        sheet.HasHeading = false;

        var factory = new ColumnIndexReaderFactory(columnIndex);
        var reader = Assert.IsType<ColumnIndexReader>(factory.GetCellReader(sheet));
        Assert.Equal(columnIndex, reader.ColumnIndex);
        Assert.NotSame(reader, factory.GetCellReader(sheet));
        Assert.Null(sheet.Heading);
    }

    [Fact]
    public void GetReader_NullSheet_ThrowsArgumentNullException()
    {
        var factory = new ColumnIndexReaderFactory(0);
        Assert.Throws<ArgumentNullException>("sheet", () => factory.GetCellReader(null!));
    }

#pragma warning disable CS0184 // The is operator is being used to test interface implementation
    [Fact]
    public void Interfaces_IColumnNameProviderCellReaderFactory_DoesNotImplement()
    {
        var factory = new ColumnIndexReaderFactory(0);
        Assert.False(factory is IColumnNameProviderCellReaderFactory);
    }

    [Fact]
    public void Interfaces_IColumnIndexProviderCellReaderFactory_DoesNotImplement()
    {
        var factory = new ColumnIndexReaderFactory(0);
        Assert.True(factory is IColumnIndexProviderCellReaderFactory);
    }

    [Fact]
    public void Interfaces_IColumnNamesProviderCellReaderFactory_DoesImplement()
    {
        var factory = new ColumnIndexReaderFactory(0);
        Assert.False(factory is IColumnNamesProviderCellReaderFactory);
    }

    [Fact]
    public void Interfaces_IColumnIndicesProviderCellReaderFactory_DoesImplement()
    {
        var factory = new ColumnIndexReaderFactory(0);
        Assert.False(factory is IColumnIndicesProviderCellReaderFactory);
    }
#pragma warning restore CS0184

    [Fact]
    public void GetColumnIndex_Invoke_ReturnsExpected()
    {
        var factory = new ColumnIndexReaderFactory(5);
        Assert.Equal(5, factory.GetColumnIndex(null!));
    }
}

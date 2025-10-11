using System;
using System.Collections.Generic;
using ExcelMapper.Tests;
using Xunit;

namespace ExcelMapper.Readers.Tests;

public class ColumnNamesReaderFactoryTests
{
    public static IEnumerable<object[]> Ctor_ParamsString()
    {
        yield return new object[] { new string[] { "ColumnName1" } };
        yield return new object[] { new string[] { "ColumnName1", "ColumnName2" } };
        yield return new object[] { new string[] { " " } };
        yield return new object[] { new string[] { "ColumnName", "ColumnName" } };
    }

    [Theory]
    [MemberData(nameof(Ctor_ParamsString))]
    public void Ctor_ColumnNames(string[] columnNames)
    {
        var reader = new ColumnNamesReaderFactory(columnNames);
        Assert.Same(columnNames, reader.ColumnNames);
    }

    [Fact]
    public void Ctor_NullColumnNames_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("columnNames", () => new ColumnNamesReaderFactory(null!));
    }

    [Fact]
    public void Ctor_EmptyColumnNames_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>("columnNames", () => new ColumnNamesReaderFactory([]));
    }

    [Fact]
    public void Ctor_NullValueInColumnNames_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>("columnNames", () => new ColumnNamesReaderFactory([null!]));
    }
    

    [Fact]
    public void GetCellReader_InvokeColumnNamesSheetWithHeadingMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();
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
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var factory = new ColumnNamesReaderFactory("NoSuchColumn");
        Assert.Null(factory.GetCellReader(sheet));
    }

    [Fact]
    public void GetCellReader_NullSheet_ThrowsArgumentNullException()
    {
        var factory = new ColumnNamesReaderFactory("Value");
        Assert.Throws<ArgumentNullException>(() => factory.GetCellReader(null!));
    }

    [Fact]
    public void GetCellReader_InvokeColumnNamesSheetNoHeadingHasHeading_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();

        var factory = new ColumnNamesReaderFactory("ColumnName");
        Assert.Throws<ExcelMappingException>(() => factory.GetCellReader(sheet));
        Assert.Null(sheet.Heading);
    }

    [Fact]
    public void GetCellReader_InvokeColumnNamesSheetNoHeadingHasNoHeading_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();
        sheet.HasHeading = false;

        var factory = new ColumnNamesReaderFactory("ColumnName");
        Assert.Throws<ExcelMappingException>(() => factory.GetCellReader(sheet));
        Assert.Null(sheet.Heading);
    }

    public static IEnumerable<object[]> GetCellsReader_TestData()
    {
        yield return new object[] { new string[] { "Value" }, new int[] { 0 } };
        yield return new object[] { new string[] { "Value", "Value" }, new int[] { 0, 0 } };
        yield return new object[] { new string[] { "vAlUE" }, new int[] { 0 } };
    }

    [Theory]
    [MemberData(nameof(GetCellsReader_TestData))]
    public void GetCellsReader_InvokeSheetWithHeading_ReturnsExpected(string[] columnNames, int[] expectedColumnIndices)
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var factory = new ColumnNamesReaderFactory(columnNames);
        var reader = Assert.IsType<ColumnIndicesReader>(factory.GetCellsReader(sheet));
        Assert.Equal(expectedColumnIndices, reader.ColumnIndices);
        Assert.NotSame(reader, factory.GetCellsReader(sheet));
    }

    public static IEnumerable<object[]> GetCellsReader_NoSuchColumn_TestData()
    {
        yield return new object[] { new string[] { "NoSuchColumn" } };
        yield return new object[] { new string[] { "Value", "NoSuchColumn" } };
    }

    [Theory]
    [MemberData(nameof(GetCellsReader_NoSuchColumn_TestData))]
    public void GetCellsReader_InvokeNoMatch_ReturnsNull(string[] columnNames)
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var factory = new ColumnNamesReaderFactory(columnNames);
        Assert.Null(factory.GetCellsReader(sheet));
    }

    [Fact]
    public void GetCellsReader_NullSheet_ThrowsArgumentNullException()
    {
        var factory = new ColumnNamesReaderFactory("Value");
        Assert.Throws<ArgumentNullException>(() => factory.GetCellsReader(null!));
    }

    [Fact]
    public void GetCellsReader_InvokeSheetNoHeadingHasHeading_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();

        var factory = new ColumnNamesReaderFactory("Value");
        Assert.Throws<ExcelMappingException>(() => factory.GetCellsReader(sheet));
        Assert.Null(sheet.Heading);
    }

    [Fact]
    public void GetCellsReader_InvokeSheetNoHeadingHasNoHeading_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();
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
        Assert.Equal(new string[] { "Column1", "Column2" }, factory.GetColumnNames(null!));
    }
}

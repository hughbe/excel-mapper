using System;
using System.Collections.Generic;
using ExcelMapper.Tests;
using Xunit;

namespace ExcelMapper.Readers.Tests;

public class ColumnNamesReaderFactoryTests
{
    [Fact]
    public void Ctor_ColumnNames()
    {
        var columnNames = new string[] { "ColumnName1", "ColumnName2" };
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

    public static IEnumerable<object[]> GetReader_TestData()
    {
        yield return new object[] { new string[] { "Value" }, new int[] { 0 } };
        yield return new object[] { new string[] { "Value", "Value" }, new int[] { 0, 0 } };
        yield return new object[] { new string[] { "vAlUE" }, new int[] { 0 } };
    }

    [Theory]
    [MemberData(nameof(GetReader_TestData))]
    public void GetReader_InvokeSheetWithHeading_ReturnsExpected(string[] columnNames, int[] expectedColumnIndices)
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var factory = new ColumnNamesReaderFactory(columnNames);
        var reader = Assert.IsType<ColumnIndicesReader>(factory.GetReader(sheet));
        Assert.Equal(expectedColumnIndices, reader.ColumnIndices);
        Assert.NotSame(reader, factory.GetReader(sheet));
    }

    public static IEnumerable<object[]> GetReader_NoSuchColumn_TestData()
    {
        yield return new object[] { new string[] { "NoSuchColumn" } };
        yield return new object[] { new string[] { "Value", "NoSuchColumn" } };
    }

    [Theory]
    [MemberData(nameof(GetReader_NoSuchColumn_TestData))]
    public void GetReader_InvokeNoMatch_ReturnsNull(string[] columnNames)
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var factory = new ColumnNamesReaderFactory(columnNames);
        Assert.Null(factory.GetReader(sheet));
    }

    [Fact]
    public void GetReader_NullSheet_ThrowsArgumentNullException()
    {
        var factory = new ColumnNamesReaderFactory("Value");
        Assert.Throws<ArgumentNullException>(() => factory.GetReader(null!));
    }

    [Fact]
    public void GetReader_InvokeSheetNoHeadingHasHeading_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();

        var factory = new ColumnNamesReaderFactory("Value");
        Assert.Throws<ExcelMappingException>(() => factory.GetReader(sheet));
        Assert.Null(sheet.Heading);
    }

    [Fact]
    public void GetReader_InvokeSheetNoHeadingHasNoHeading_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();
        sheet.HasHeading = false;

        var factory = new ColumnNamesReaderFactory("Value");
        Assert.Throws<ExcelMappingException>(() => factory.GetReader(sheet));
        Assert.Null(sheet.Heading);
    }
}

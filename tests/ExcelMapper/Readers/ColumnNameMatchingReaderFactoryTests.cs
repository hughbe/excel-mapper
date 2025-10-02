using System;
using System.Collections.Generic;
using ExcelMapper.Tests;
using Xunit;

namespace ExcelMapper.Readers.Tests;

public class ColumnNameMatchingReaderFactoryTests
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
        var factory = new ColumnNameMatchingReaderFactory(columnNames);
        Assert.Same(columnNames, factory.ColumnNames);
        Assert.Null(factory.Predicate);
    }

    [Fact]
    public void Ctor_NullColumnNames_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("columnNames", () => new ColumnNameMatchingReaderFactory((string[])null!));
    }

    [Fact]
    public void Ctor_EmptyColumnNames_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>("columnNames", () => new ColumnNameMatchingReaderFactory([]));
    }

    [Fact]
    public void Ctor_NullValueInColumnNames_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>("columnNames", () => new ColumnNameMatchingReaderFactory([null!]));
    }

    [Fact]
    public void Ctor_Predicate()
    {
        Func<string, bool> predicate = columnName => true; 
        var factory = new ColumnNameMatchingReaderFactory(predicate);
        Assert.Null(factory.ColumnNames);
        Assert.Same(predicate, factory.Predicate);
    }

    [Fact]
    public void Ctor_NullPredicate_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("predicate", () => new ColumnNameMatchingReaderFactory((Func<string, bool>)null!));
    }

    [Fact]
    public void GetReader_InvokeColumnNamesSheetWithHeadingMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var factory = new ColumnNameMatchingReaderFactory("NoSuchColumn", "Value");
        var reader = Assert.IsType<ColumnIndexReader>(factory.GetReader(sheet));
        Assert.Equal(0, reader.ColumnIndex);
        Assert.NotSame(reader, factory.GetReader(sheet));
    }

    [Fact]
    public void GetReader_InvokeColumnNamesNoMatch_ReturnsNull()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var factory = new ColumnNameMatchingReaderFactory("NoSuchColumn");
        Assert.Null(factory.GetReader(sheet));
    }

    [Fact]
    public void GetReader_InvokeColumnNamesSheetNoHeadingHasHeading_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();

        var factory = new ColumnNameMatchingReaderFactory("ColumnName");
        Assert.Throws<ExcelMappingException>(() => factory.GetReader(sheet));
        Assert.Null(sheet.Heading);
    }

    [Fact]
    public void GetReader_InvokeColumnNamesSheetNoHeadingHasNoHeading_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();
        sheet.HasHeading = false;

        var factory = new ColumnNameMatchingReaderFactory("ColumnName");
        Assert.Throws<ExcelMappingException>(() => factory.GetReader(sheet));
        Assert.Null(sheet.Heading);
    }

    [Fact]
    public void GetReader_InvokePredicateSheetWithHeadingMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        List<string> calls = [];
        bool Match(string columnName)
        {
            calls.Add(columnName);
            return columnName == "Value";
        }
        var factory = new ColumnNameMatchingReaderFactory(Match);
        var reader = Assert.IsType<ColumnIndexReader>(factory.GetReader(sheet));
        Assert.Equal(0, reader.ColumnIndex);
        Assert.Equal(["Value"], calls);
        Assert.NotSame(reader, factory.GetReader(sheet));
    }

    [Fact]
    public void GetReader_InvokePredicateNoMatch_ReturnsNull()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        List<string> calls = [];
        bool Match(string columnName)
        {
            calls.Add(columnName);
            return columnName != "Value";
        }
        var factory = new ColumnNameMatchingReaderFactory(Match);
        Assert.Null(factory.GetReader(sheet));
        Assert.Equal(["Value"], calls);
    }

    [Fact]
    public void GetReader_NullSheet_ThrowsArgumentNullException()
    {
        static bool Match(string columnName) => true;
        var factory = new ColumnNameMatchingReaderFactory(Match);
        Assert.Throws<ArgumentNullException>(() => factory.GetReader(null!));
    }

    [Fact]
    public void GetReader_InvokePredicateSheetNoHeadingHasHeading_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();

        List<string> calls = [];
        bool Match(string columnName)
        {
            calls.Add(columnName);
            return columnName == "Value";
        }
        var factory = new ColumnNameMatchingReaderFactory(Match);
        Assert.Throws<ExcelMappingException>(() => factory.GetReader(sheet));
        Assert.Null(sheet.Heading);
    }

    [Fact]
    public void GetReader_InvokePredicateSheetNoHeadingHasNoHeading_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();
        sheet.HasHeading = false;

        List<string> calls = [];
        bool Match(string columnName)
        {
            calls.Add(columnName);
            return columnName == "Value";
        }
        var factory = new ColumnNameMatchingReaderFactory(Match);
        Assert.Throws<ExcelMappingException>(() => factory.GetReader(sheet));
        Assert.Null(sheet.Heading);
    }
}

using System;
using System.Collections.Generic;
using ExcelMapper.Tests;
using Xunit;

namespace ExcelMapper.Readers.Tests;

public class ColumnNameMatchingReaderFactoryTests
{
    [Fact]
    public void Ctor_ColumnName()
    {
        var factory = new ColumnNameMatchingReaderFactory(e => e == "ColumnName");
        Assert.NotNull(factory);
    }

    [Fact]
    public void Ctor_NullColumnName_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("predicate", () => new ColumnNameMatchingReaderFactory(null!));
    }

    [Fact]
    public void GetReader_InvokeSheetWithHeadingMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        List<string> calls = [];
        bool Match(string columnName)
        {
            calls.Add(columnName);
            return columnName == "Value";
        };
        var factory = new ColumnNameMatchingReaderFactory(Match);
        var reader = Assert.IsType<ColumnIndexReader>(factory.GetReader(sheet));
        Assert.Equal(0, reader.ColumnIndex);
        Assert.Equal(["Value"], calls);
        Assert.NotSame(reader, factory.GetReader(sheet));
    }

    [Fact]
    public void GetReader_InvokeNoMatch_ReturnsNull()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        List<string> calls = [];
        bool Match(string columnName)
        {
            calls.Add(columnName);
            return columnName != "Value";
        };
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
    public void GetReader_InvokeSheetNoHeadingHasHeading_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();

        List<string> calls = [];
        bool Match(string columnName)
        {
            calls.Add(columnName);
            return columnName == "Value";
        };
        var factory = new ColumnNameMatchingReaderFactory(Match);
        Assert.Throws<ExcelMappingException>(() => factory.GetReader(sheet));
        Assert.Null(sheet.Heading);
    }

    [Fact]
    public void GetReader_InvokeSheetNoHeadingHasNoHeading_ThrowsExcelMappingException()
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
        ;
        var factory = new ColumnNameMatchingReaderFactory(Match);
        Assert.Throws<ExcelMappingException>(() => factory.GetReader(sheet));
        Assert.Null(sheet.Heading);
    }
}

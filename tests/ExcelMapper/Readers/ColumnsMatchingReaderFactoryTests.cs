using System;
using System.Collections.Generic;
using ExcelMapper.Abstractions;
using ExcelMapper.Tests;
using Xunit;

namespace ExcelMapper.Readers.Tests;

public class ColumnsMatchingReaderFactoryTests
{
    [Fact]
    public void Ctor_IExcelColumnMatcher()
    {
        bool Match(string columnName) => true;
        var matcher = new PredicateColumnMatcher(Match);
        var factory = new ColumnsMatchingReaderFactory(matcher);
        Assert.Same(matcher, factory.Matcher);
    }

    [Fact]
    public void Ctor_NullMatcher_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("matcher", () => new ColumnsMatchingReaderFactory((IExcelColumnMatcher)null!));
    }

    [Fact]
    public void GetCellReader_InvokePredicateSheetWithHeadingMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("MultipleStrings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        List<string> calls = [];
        bool Match(string columnName)
        {
            calls.Add(columnName);
            return columnName.StartsWith("Value");
        }
        var factory = new ColumnsMatchingReaderFactory(new PredicateColumnMatcher(Match));
        var reader = Assert.IsType<ColumnIndexReader>(factory.GetCellReader(sheet));
        Assert.Equal(0, reader.ColumnIndex);
        Assert.Equal(["Value"], calls);
        Assert.NotSame(reader, factory.GetCellReader(sheet));
    }

    [Fact]
    public void GetCellReader_InvokePredicateNoMatch_ReturnsNull()
    {
        using var importer = Helpers.GetImporter("MultipleStrings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        List<string> calls = [];
        bool Match(string columnName)
        {
            calls.Add(columnName);
            return false;
        }
        var factory = new ColumnsMatchingReaderFactory(new PredicateColumnMatcher(Match));
        Assert.Null(factory.GetCellReader(sheet));
        Assert.Equal(["Value", "Value2", "NoSuchValue"], calls);
    }

    [Fact]
    public void GetCellReader_NullSheet_ThrowsArgumentNullException()
    {
        static bool Match(string columnName) => true;
        var factory = new ColumnsMatchingReaderFactory(new PredicateColumnMatcher(Match));
        Assert.Throws<ArgumentNullException>("sheet", () => factory.GetCellReader(null!));
    }

    [Fact]
    public void GetCellReader_InvokePredicateSheetNoHeadingHasHeading_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("MultipleStrings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();

        List<string> calls = [];
        bool Match(string columnName)
        {
            calls.Add(columnName);
            return columnName.StartsWith("Value");
        }
        var factory = new ColumnsMatchingReaderFactory(new PredicateColumnMatcher(Match));
        Assert.Throws<ExcelMappingException>(() => factory.GetCellReader(sheet));
        Assert.Null(sheet.Heading);
    }

    [Fact]
    public void GetCellReader_InvokePredicateSheetNoHeadingHasNoHeading_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("MultipleStrings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();
        sheet.HasHeading = false;

        List<string> calls = [];
        bool Match(string columnName)
        {
            calls.Add(columnName);
            return columnName.StartsWith("Value");
        }
        var factory = new ColumnsMatchingReaderFactory(new PredicateColumnMatcher(Match));
        Assert.Throws<ExcelMappingException>(() => factory.GetCellReader(sheet));
        Assert.Null(sheet.Heading);
    }

    [Fact]
    public void GetCellsReader_InvokePredicateSheetWithHeadingMatch_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("MultipleStrings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        List<string> calls = [];
        bool Match(string columnName)
        {
            calls.Add(columnName);
            return columnName.StartsWith("Value");
        }
        var factory = new ColumnsMatchingReaderFactory(new PredicateColumnMatcher(Match));
        var reader = Assert.IsType<ColumnIndicesReader>(factory.GetCellsReader(sheet));
        Assert.Equal([0, 1], reader.ColumnIndices);
        Assert.Equal(["Value", "Value2", "NoSuchValue"], calls);
        Assert.NotSame(reader, factory.GetCellsReader(sheet));
    }

    [Fact]
    public void GetCellsReader_InvokePredicateNoMatch_ReturnsNull()
    {
        using var importer = Helpers.GetImporter("MultipleStrings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        List<string> calls = [];
        bool Match(string columnName)
        {
            calls.Add(columnName);
            return false;
        }
        var factory = new ColumnsMatchingReaderFactory(new PredicateColumnMatcher(Match));
        Assert.Null(factory.GetCellsReader(sheet));
        Assert.Equal(["Value", "Value2", "NoSuchValue"], calls);
    }

    [Fact]
    public void GetCellsReader_NullSheet_ThrowsArgumentNullException()
    {
        static bool Match(string columnName) => true;
        var factory = new ColumnsMatchingReaderFactory(new PredicateColumnMatcher(Match));
        Assert.Throws<ArgumentNullException>("sheet", () => factory.GetCellsReader(null!));
    }

    [Fact]
    public void GetCellsReader_InvokePredicateSheetNoHeadingHasHeading_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("MultipleStrings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();

        List<string> calls = [];
        bool Match(string columnName)
        {
            calls.Add(columnName);
            return columnName.StartsWith("Value");
        }
        var factory = new ColumnsMatchingReaderFactory(new PredicateColumnMatcher(Match));
        Assert.Throws<ExcelMappingException>(() => factory.GetCellsReader(sheet));
        Assert.Null(sheet.Heading);
    }

    [Fact]
    public void GetCellsReader_InvokePredicateSheetNoHeadingHasNoHeading_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("MultipleStrings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();
        sheet.HasHeading = false;

        List<string> calls = [];
        bool Match(string columnName)
        {
            calls.Add(columnName);
            return columnName.StartsWith("Value");
        }
        var factory = new ColumnsMatchingReaderFactory(new PredicateColumnMatcher(Match));
        Assert.Throws<ExcelMappingException>(() => factory.GetCellsReader(sheet));
        Assert.Null(sheet.Heading);
    }

#pragma warning disable CS0184 // The is operator is being used to test interface implementation
    [Fact]
    public void Interfaces_IColumnNameProviderCellReaderFactory_DoesNotImplement()
    {
        static bool Match(string columnName) => true;
        var factory = new ColumnsMatchingReaderFactory(new PredicateColumnMatcher(Match));
        Assert.False(factory is IColumnNameProviderCellReaderFactory);
    }

    [Fact]
    public void Interfaces_IColumnIndexProviderCellReaderFactory_DoesNotImplement()
    {
        static bool Match(string columnName) => true;
        var factory = new ColumnsMatchingReaderFactory(new PredicateColumnMatcher(Match));
        Assert.False(factory is IColumnIndexProviderCellReaderFactory);
    }

    [Fact]
    public void Interfaces_IColumnNamesProviderCellReaderFactory_DoesImplement()
    {
        static bool Match(string columnName) => true;
        var factory = new ColumnsMatchingReaderFactory(new PredicateColumnMatcher(Match));
        Assert.True(factory is IColumnNamesProviderCellReaderFactory);
    }

    [Fact]
    public void Interfaces_IColumnIndicesProviderCellReaderFactory_DoesImplement()
    {
        static bool Match(string columnName) => true;
        var factory = new ColumnsMatchingReaderFactory(new PredicateColumnMatcher(Match));
        Assert.True(factory is IColumnIndicesProviderCellReaderFactory);
    }
#pragma warning restore CS0184

    [Fact]
    public void GetColumnNames_Invoke_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("MultipleStrings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        static bool Match(string columnName) => columnName.StartsWith("Value");
        var factory = new ColumnsMatchingReaderFactory(new PredicateColumnMatcher(Match));
        Assert.Equal(new string[] { "Value", "Value2" }, factory.GetColumnNames(sheet));
    }

    [Fact]
    public void GetColumnNames_InvokeNoHeading_ReturnsNull()
    {
        using var importer = Helpers.GetImporter("MultipleStrings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();

        static bool Match(int columnIndex) => columnIndex == 0 || columnIndex == 1;
        var factory = new ColumnsMatchingReaderFactory(new CustomColumnMatcher(Match));
        Assert.Null(factory.GetColumnNames(sheet));
    }

    [Fact]
    public void GetColumnNames_NullSheet_ThrowsArgumentNullException()
    {
        static bool Match(string columnName) => true;
        var factory = new ColumnsMatchingReaderFactory(new PredicateColumnMatcher(Match));
        Assert.Throws<ArgumentNullException>("sheet", () => factory.GetColumnNames(null!));
    }
    [Fact]
    public void GetColumnIndices_Invoke_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("MultipleStrings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();
        sheet.ReadHeading();

        static bool Match(string columnName) => columnName.StartsWith("Value");
        var factory = new ColumnsMatchingReaderFactory(new PredicateColumnMatcher(Match));
        Assert.Equal(new int[] { 0, 1 }, factory.GetColumnIndices(sheet));
    }

    [Fact]
    public void GetColumnIndices_InvokeNoHeading_ReturnsExpeced()
    {
        using var importer = Helpers.GetImporter("MultipleStrings.xlsx");
        ExcelSheet sheet = importer.ReadSheet();

        static bool Match(int columnIndex) => columnIndex == 0 || columnIndex == 1;
        var factory = new ColumnsMatchingReaderFactory(new CustomColumnMatcher(Match));
        Assert.Equal(new int[] { 0, 1 }, factory.GetColumnIndices(sheet));
    }

    [Fact]
    public void GetColumnIndices_NullSheet_ThrowsArgumentNullException()
    {
        static bool Match(string columnName) => true;
        var factory = new ColumnsMatchingReaderFactory(new PredicateColumnMatcher(Match));
        Assert.Throws<ArgumentNullException>("sheet", () => factory.GetColumnIndices(null!));
    }

    private class CustomColumnMatcher : IExcelColumnMatcher
    {
        private readonly Func<int, bool> _match;

        public CustomColumnMatcher(Func<int, bool> match)
        {
            ArgumentNullException.ThrowIfNull(match);
            _match = match;
        }

        public bool ColumnMatches(ExcelSheet sheet, int columnIndex) => _match(columnIndex);
    }
}

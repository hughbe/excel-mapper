using System;
using System.Collections.Generic;
using ExcelMapper.Tests;
using Xunit;

namespace ExcelMapper.Readers.Tests;

public class PredicateColumnMatcherTests
{
    [Fact]
    public void Ctor_FuncStringBool()
    {
        Func<string, bool> predicate = (columnName) => true;
        var matcher = new PredicateColumnMatcher(predicate);
        Assert.Same(predicate, matcher.Predicate);
    }

    [Fact]
    public void Ctor_NullPredicate_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("predicate", () => new PredicateColumnMatcher(null!));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ColumnMatches_Invoke_ReturnsExpected(bool result)
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        List<string> calls = [];
        bool Predicate(string columnName)
        {
            calls.Add(columnName);
            return result;
        }
        var matcher = new PredicateColumnMatcher(Predicate);
        Assert.Equal(result, matcher.ColumnMatches(sheet, 0));
        Assert.Equal(["Value"], calls);
    }

    [Fact]
    public void ColumnMatches_NullSheet_ThrowsArgumentNullException()
    {
        bool Predicate(string columnName) => true;
        var matcher = new PredicateColumnMatcher(Predicate);
        Assert.Throws<ArgumentNullException>("sheet", () => matcher.ColumnMatches(null!, 0));
    }

    [Fact]
    public void ColumnMatches_SheetWithNoHeadingHasHeading_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();

        bool Predicate(string ColumnName) => true;
        var matcher = new PredicateColumnMatcher(Predicate);
        Assert.Throws<ExcelMappingException>(() => matcher.ColumnMatches(sheet, 0));
        Assert.Null(sheet.Heading);
    }

    [Fact]
    public void ColumnMatches_SheetWithNoHeadingHasNoHeading_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;

        bool Predicate(string ColumnName) => true;
        var matcher = new PredicateColumnMatcher(Predicate);
        Assert.Throws<ExcelMappingException>(() => matcher.ColumnMatches(sheet, 0));
        Assert.Null(sheet.Heading);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(123)]
    public void ColumnMatches_InvalidColumnIndex_ThrowsArgumentOutOfRangeException(int columnIndex)
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        bool Predicate(string ColumnName) => true;
        var matcher = new PredicateColumnMatcher(Predicate);
        Assert.Throws<ArgumentOutOfRangeException>("columnIndex", () => matcher.ColumnMatches(sheet, columnIndex));
    }
}

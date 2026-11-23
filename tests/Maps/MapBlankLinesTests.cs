using System.Linq;

namespace ExcelMapper.Tests;

/// <summary>
/// Blank rows in the middle of a sheet. The CSV reader's row count leaves out empty lines, so
/// these check that rows after them are still read, and that trailing empty lines are ignored.
/// </summary>
public class MapBlankLinesTests
{
    [Theory]
    [InlineData("Name\na\n\nb\n")]
    [InlineData("Name\r\na\r\n\r\nb\r\n")]
    [InlineData("Name\na\n \nb\n")]
    [InlineData("Name\na\n\nb")]
    public void ReadRows_CsvEmptyLine_ReadsRowsAfterIt(string csv)
    {
        using var importer = Helpers.GetCsvImporter(csv);
        var rows = importer.ReadSheet().ReadRows<StringRow>().ToArray();
        Assert.Equal(["a", null, "b"], rows.Select(r => r.Name));
    }

    [Fact]
    public void ReadRows_CsvEmptyLineSkipBlankLines_ReadsRowsAfterIt()
    {
        using var importer = Helpers.GetCsvImporter("Name\na\n\nb\n");
        importer.Configuration.SkipBlankLines = true;
        var rows = importer.ReadSheet().ReadRows<StringRow>().ToArray();
        Assert.Equal(["a", "b"], rows.Select(r => r.Name));
    }

    [Fact]
    public void ReadRows_CsvSeveralEmptyLines_ReadsRowsAfterThem()
    {
        using var importer = Helpers.GetCsvImporter("Name\na\n\n\n\nb\n\nc\n");
        importer.Configuration.SkipBlankLines = true;
        var rows = importer.ReadSheet().ReadRows<StringRow>().ToArray();
        Assert.Equal(["a", "b", "c"], rows.Select(r => r.Name));
    }

    [Fact]
    public void ReadRows_CsvEmptyLineMultipleColumns_ReadsRowsAfterIt()
    {
        using var importer = Helpers.GetCsvImporter("Name,Age\na,1\n\nb,2\n,\nc,3\n");
        importer.Configuration.SkipBlankLines = true;
        var rows = importer.ReadSheet().ReadRows<PersonRow>().ToArray();
        Assert.Equal(["a", "b", "c"], rows.Select(r => r.Name));
        Assert.Equal([1, 2, 3], rows.Select(r => r.Age));
    }

    [Fact]
    public void ReadRows_CsvEmptyLineMultipleColumnsNotSkipped_MapsBlankRow()
    {
        using var importer = Helpers.GetCsvImporter("Name,Age\na,1\n\nb,2\n");
        var rows = importer.ReadSheet().ReadRows<NullablePersonRow>().ToArray();
        Assert.Equal(["a", null, "b"], rows.Select(r => r.Name));
        Assert.Equal([1, null, 2], rows.Select(r => r.Age));
    }

    [Theory]
    [InlineData("Name\na\n\n")]
    [InlineData("Name\na\n\n\n\n")]
    [InlineData("Name\r\na\r\n\r\n\r\n")]
    public void ReadRows_CsvTrailingEmptyLines_Ignored(string csv)
    {
        using var importer = Helpers.GetCsvImporter(csv);
        var rows = importer.ReadSheet().ReadRows<StringRow>().ToArray();
        Assert.Equal(["a"], rows.Select(r => r.Name));
    }

    [Fact]
    public void ReadRows_CsvTrailingEmptyLinesMultipleColumns_Ignored()
    {
        // Mapping a trailing empty line would fail, as Age isn't nullable.
        using var importer = Helpers.GetCsvImporter("Name,Age\na,1\n\n\n");
        var rows = importer.ReadSheet().ReadRows<PersonRow>().ToArray();
        Assert.Equal(["a"], rows.Select(r => r.Name));
    }

    [Fact]
    public void ReadRows_CsvEmptyLinesInMiddleAndEnd_IgnoresTrailingOnly()
    {
        using var importer = Helpers.GetCsvImporter("Name\na\n\nb\n\n\n");
        var rows = importer.ReadSheet().ReadRows<StringRow>().ToArray();
        Assert.Equal(["a", null, "b"], rows.Select(r => r.Name));
    }

    [Fact]
    public void ReadRows_CsvEmptyLinesAboveHeading_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("\n\nName\na\nb\n");
        var sheet = importer.ReadSheet();
        sheet.HeadingIndex = 2;
        var rows = sheet.ReadRows<StringRow>().ToArray();
        Assert.Equal(["a", "b"], rows.Select(r => r.Name));
        Assert.Equal(4, sheet.CurrentRowIndex);
    }

    [Fact]
    public void ReadRows_CsvEmptyLinesNoHeading_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("a\n\nb\n");
        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;
        importer.Configuration.SkipBlankLines = true;
        var rows = sheet.ReadRows<ColumnIndexRow>().ToArray();
        Assert.Equal(["a", "b"], rows.Select(r => r.Value));
    }

    [Fact]
    public void ReadRow_CsvEmptyLine_ReadsRowsAfterIt()
    {
        using var importer = Helpers.GetCsvImporter("Name\na\n\nb\n\n");
        importer.Configuration.SkipBlankLines = true;
        var sheet = importer.ReadSheet();
        Assert.Equal("a", sheet.ReadRow<StringRow>().Name);
        Assert.Equal("b", sheet.ReadRow<StringRow>().Name);
        Assert.False(sheet.TryReadRow<StringRow>(out _));
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<StringRow>());
    }

    [Fact]
    public void ReadRows_CsvStartIndexAfterEmptyLine_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Name\na\n\nb\nc\n");
        var rows = importer.ReadSheet().ReadRows<StringRow>(3, 2).ToArray();
        Assert.Equal(["b", "c"], rows.Select(r => r.Name));
    }

    [Fact]
    public void ReadRows_CsvDataRangePastEnd_DoesNotRepeatLastRow()
    {
        using var importer = Helpers.GetCsvImporter("Name\na\n\nb\n");
        var sheet = importer.ReadSheet();
        sheet.DataRange = ExcelRange.Parse("1:10");
        var rows = sheet.ReadRows<StringRow>().ToArray();
        Assert.Equal(["a", null, "b"], rows.Select(r => r.Name));
        Assert.Equal(3, sheet.CurrentRowIndex);
    }

    [Fact]
    public void ReadRows_CsvEmptyLineDataRangeFromEnd_ExcludesLastRows()
    {
        // For example, leave out a totals row at the bottom.
        using var importer = Helpers.GetCsvImporter("Name\na\n\nb\nTotal\n");
        var sheet = importer.ReadSheet();
        sheet.DataRange = new ExcelRange(..^1, Range.All);
        var rows = sheet.ReadRows<StringRow>().ToArray();
        Assert.Equal(["a", null, "b"], rows.Select(r => r.Name));
        Assert.Equal(3, sheet.CurrentRowIndex);
    }

    [Fact]
    public void ReadRows_CsvEmptyLineDataRangeFromEndSkipBlankLines_ExcludesLastRows()
    {
        using var importer = Helpers.GetCsvImporter("Name\na\n\n\nb\nc\nTotal\n");
        importer.Configuration.SkipBlankLines = true;
        var sheet = importer.ReadSheet();
        sheet.DataRange = new ExcelRange(..^2, Range.All);
        var rows = sheet.ReadRows<StringRow>().ToArray();
        Assert.Equal(["a", "b"], rows.Select(r => r.Name));
    }

    [Fact]
    public void ReadRows_CsvTrailingEmptyLinesDataRangeFromEnd_IgnoresTrailingEmptyLines()
    {
        using var importer = Helpers.GetCsvImporter("Name\na\nb\nTotal\n\n\n");
        var sheet = importer.ReadSheet();
        sheet.DataRange = new ExcelRange(..^1, Range.All);
        var rows = sheet.ReadRows<StringRow>().ToArray();
        Assert.Equal(["a", "b"], rows.Select(r => r.Name));
    }

    [Fact]
    public void ReadRows_CsvEmptyLineDataRangeStartFromEnd_ReturnsLastRows()
    {
        using var importer = Helpers.GetCsvImporter("a\n\nb\nc\n");
        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;
        sheet.DataRange = new ExcelRange(^2.., Range.All);
        var rows = sheet.ReadRows<ColumnIndexRow>().ToArray();
        Assert.Equal(["b", "c"], rows.Select(r => r.Value));
    }

    [Fact]
    public void ReadRows_CsvEmptyLineHeadingFromEnd_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Report\n\nName\na\nb\n");
        var sheet = importer.ReadSheet();
        sheet.DataRange = new ExcelRange(^3.., Range.All);
        var rows = sheet.ReadRows<StringRow>().ToArray();
        Assert.Equal(["a", "b"], rows.Select(r => r.Name));
        Assert.Equal(["Name"], sheet.Heading!.ColumnNames);
    }

    [Fact]
    public void ReadRow_CsvEmptyLineDataRangeFromEnd_ReturnsExpected()
    {
        // Counting the rows goes back to the current row afterwards.
        using var importer = Helpers.GetCsvImporter("Name\na\n\nb\nTotal\n");
        var sheet = importer.ReadSheet();
        sheet.DataRange = new ExcelRange(..^1, Range.All);
        Assert.Equal("a", sheet.ReadRow<StringRow>().Name);
        Assert.Null(sheet.ReadRow<StringRow>().Name);
        Assert.Equal("b", sheet.ReadRow<StringRow>().Name);
        Assert.False(sheet.TryReadRow<StringRow>(out _));
    }

    [Fact]
    public void ReadRows_XlsxBlankRowsDataRangeFromEnd_ExcludesLastRows()
    {
        using var importer = XlsxBuilder.Create(["Name"], ["a"], [null], ["b"], ["Total"]);
        var sheet = importer.ReadSheet();
        sheet.DataRange = new ExcelRange(..^1, Range.All);
        var rows = sheet.ReadRows<StringRow>().ToArray();
        Assert.Equal(["a", null, "b"], rows.Select(r => r.Name));
    }

    [Fact]
    public void ReadRows_CsvQuotedValueWithEmptyLine_ReturnsExpected()
    {
        using var importer = Helpers.GetCsvImporter("Name\n\"x\n\ny\"\nb\n");
        var rows = importer.ReadSheet().ReadRows<StringRow>().ToArray();
        Assert.Equal(["x\n\ny", "b"], rows.Select(r => r.Name));
    }

    [Fact]
    public void ReadRows_XlsxBlankRows_ReturnsExpected()
    {
        // A blank row between rows is mapped; blank rows at the end aren't rows.
        using var importer = XlsxBuilder.Create(["Name"], ["a"], [null], ["b"], [null], [null]);
        var rows = importer.ReadSheet().ReadRows<StringRow>().ToArray();
        Assert.Equal(["a", null, "b"], rows.Select(r => r.Name));
    }

    [Fact]
    public void ReadRows_XlsxBlankRowsBetweenRows_MapsEachBlankRow()
    {
        using var importer = XlsxBuilder.Create(["Name", "Age"], ["a", 1], [null, null], [null, null], [null, null], ["b", 2], [null, null]);
        importer.Configuration.RegisterClassMap<PersonRow>(c =>
        {
            c.Map(r => r.Name);
            c.Map(r => r.Age).WithEmptyFallback(-1);
        });
        var rows = importer.ReadSheet().ReadRows<PersonRow>().ToArray();
        Assert.Equal(["a", null, null, null, "b"], rows.Select(r => r.Name));
        Assert.Equal([1, -1, -1, -1, 2], rows.Select(r => r.Age));
    }

    [Fact]
    public void ReadRows_XlsxOnlyBlankRowsAfterHeading_ReturnsNoRows()
    {
        using var importer = XlsxBuilder.Create(["Name"], [null], [null]);
        Assert.Empty(importer.ReadSheet().ReadRows<StringRow>());
    }

    [Fact]
    public void ReadRows_XlsxTrailingBlankRows_NotCountedFromEnd()
    {
        // Rows counted from the end leave out the blank rows at the end.
        using var importer = XlsxBuilder.Create(["Name"], ["a"], [null], ["b"], [null], [null]);
        var sheet = importer.ReadSheet();
        sheet.DataRange = new ExcelRange(..^1, Range.All);
        Assert.Equal(["a", null], sheet.ReadRows<StringRow>().Select(r => r.Name));
    }

    [Fact]
    public void ReadRows_XlsxBlankRowsBetweenRowsPreserveFormatting_AreEmpty()
    {
        using var importer = XlsxBuilder.Create(["Name"], ["a"], [null], ["b"]);
        importer.Configuration.RegisterClassMap<StringRow>(c => c.Map(r => r.Name).MakePreserveFormatting());
        Assert.Equal(["a", null, "b"], importer.ReadSheet().ReadRows<StringRow>().Select(r => r.Name));
    }

    [Fact]
    public void ReadRows_XlsxBlankRowsBetweenRowsWithRowNumbers_NumberEveryRow()
    {
        using var importer = XlsxBuilder.Create(["Name"], ["a"], [null], [null], ["b"], [null]);
        importer.Configuration.RegisterClassMap<NumberedRow>(c =>
        {
            c.Map(r => r.Name);
            c.MapRowNumber(r => r.Row);
        });
        var rows = importer.ReadSheet().ReadRows<NumberedRow>().ToArray();
        Assert.Equal([2, 3, 4, 5], rows.Select(r => r.Row));
        Assert.Equal(["a", null, null, "b"], rows.Select(r => r.Name));
    }

    public class NumberedRow
    {
        public string? Name { get; set; }
        public int Row { get; set; }
    }

    [Fact]
    public void ReadRows_XlsxBlankRowsSkipBlankLines_ReturnsExpected()
    {
        using var importer = XlsxBuilder.Create(["Name"], ["a"], [null], ["b"], [null]);
        importer.Configuration.SkipBlankLines = true;
        var rows = importer.ReadSheet().ReadRows<StringRow>().ToArray();
        Assert.Equal(["a", "b"], rows.Select(r => r.Name));
    }

    public class StringRow
    {
        public string? Name { get; set; }
    }

    public class PersonRow
    {
        public string? Name { get; set; }
        public int Age { get; set; }
    }

    public class NullablePersonRow
    {
        public string? Name { get; set; }
        public int? Age { get; set; }
    }

    public class ColumnIndexRow
    {
        [ExcelColumnIndex(0)]
        public string? Value { get; set; }
    }
}

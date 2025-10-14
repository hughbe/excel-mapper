using System.Collections.Generic;
using ExcelMapper.Abstractions;
using Xunit;

namespace ExcelMapper.Tests;

public class CantReadTests
{
    [Fact]
    public void CantRead_NoSuchColumnIndex_ThrowsCorrectException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoSuchColumnIndexClass>());
        Assert.StartsWith($"Could not read value for member \"Member\" for column at index {int.MaxValue}", ex.Message);
        Assert.Equal(0, ex.RowIndex);
        Assert.Equal(-1, ex.ColumnIndex);
    }

    private class NoSuchColumnIndexClass
    {
        [ExcelColumnIndex(int.MaxValue)]
        public string Member { get; set; } = default!;
    }

    [Fact]
    public void CantRead_NoSuchColumnName_ThrowsCorrectException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoSuchColumnNameClass>());
        Assert.StartsWith("Could not read value for member \"Member\" for column \"NoSuchColumn\"", ex.Message);
        Assert.Equal(0, ex.RowIndex);
        Assert.Equal(-1, ex.ColumnIndex);
    }

    private class NoSuchColumnNameClass
    {
        [ExcelColumnName("NoSuchColumn")]
        public string Member { get; set; } = default!;
    }

    [Fact]
    public void CantRead_NoSuchColumnMatching_ThrowsCorrectException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoSuchColumnMatchingClass>());
        Assert.StartsWith($"Could not read value for member \"Member\" (no columns matching)", ex.Message);
        Assert.Equal(0, ex.RowIndex);
        Assert.Equal(-1, ex.ColumnIndex);
    }

    [Fact]
    public void CantRead_NoSuchColumnMatchingNoHeading_ThrowsCorrectException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();

        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoSuchColumnMatchingClass>());
        Assert.StartsWith($"Could not read value for member \"Member\" (no columns matching)", ex.Message);
        Assert.Equal(0, ex.RowIndex);
        Assert.Equal(-1, ex.ColumnIndex);
    }

    private class NoSuchColumnMatchingClass
    {
        [ExcelColumnMatching(typeof(NeverMatcher))]
        public string Member { get; set; } = default!;
    }

    private class NeverMatcher : IExcelColumnMatcher
    {
        public bool ColumnMatches(ExcelSheet sheet, int columnIndex) => false;
    }

    [Fact]
    public void CantRead_SplitNoSuchColumnIndex_ThrowsCorrectException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoSuchSplitColumnIndexClass>());
        Assert.StartsWith($"Could not read value for member \"Member\" for column at index {int.MaxValue}", ex.Message);
        Assert.Equal(0, ex.RowIndex);
        Assert.Equal(-1, ex.ColumnIndex);
    }

    private class NoSuchSplitColumnIndexClass
    {
        [ExcelColumnIndex(int.MaxValue)]
        public IEnumerable<string> Member { get; set; } = default!;
    }

    [Fact]
    public void CantRead_SplitNoSuchColumnName_ThrowsCorrectException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoSuchSplitColumnNameClass>());
        Assert.StartsWith($"Could not read value for member \"Member\" for column \"NoSuchColumn\"", ex.Message);
        Assert.Equal(0, ex.RowIndex);
        Assert.Equal(-1, ex.ColumnIndex);
    }

    private class NoSuchSplitColumnNameClass
    {
        [ExcelColumnName("NoSuchColumn")]
        public IEnumerable<string> Member { get; set; } = default!;
    }

    [Fact]
    public void CantRead_EnumerableNoSuchColumnIndex_ThrowsCorrectException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoSuchEnumerableColumnIndexClass>());
        Assert.StartsWith($"Could not read value for member \"Member\" for columns at indices 1, {int.MaxValue}", ex.Message);
        Assert.Equal(0, ex.RowIndex);
        Assert.Equal(-1, ex.ColumnIndex);
    }

    private class NoSuchEnumerableColumnIndexClass
    {
        [ExcelColumnIndices(1, int.MaxValue)]
        public IEnumerable<string> Member { get; set; } = default!;
    }

    [Fact]
    public void CantRead_EnumerableNoSuchColumnName_ThrowsCorrectException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NoSuchEnumerableColumnNameClass>());
        Assert.StartsWith($"Could not read value for member \"Member\" for columns \"Value\", \"NoSuchColumn\"", ex.Message);
        Assert.Equal(0, ex.RowIndex);
        Assert.Equal(-1, ex.ColumnIndex);
    }

    private class NoSuchEnumerableColumnNameClass
    {
        [ExcelColumnNames("Value", "NoSuchColumn")]
        public IEnumerable<string> Member { get; set; } = default!;
    }
}

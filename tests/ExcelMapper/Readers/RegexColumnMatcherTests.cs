using System.Text.RegularExpressions;
using ExcelMapper.Tests;

namespace ExcelMapper.Readers.Tests;

public class RegexColumnMatcherTests
{
    [Fact]
    public void Ctor_Regex()
    {
        var regex = new Regex("Regex");
        var matcher = new RegexColumnMatcher(regex);
        Assert.Same(regex, matcher.Regex);
    }

    [Fact]
    public void Ctor_NullRegex_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("regex", () => new RegexColumnMatcher(null!));
    }

    [Theory]
    [InlineData("Value", true)]
    [InlineData("NoSuchColumn", false)]
    public void ColumnMatches_Invoke_ReturnsExpected(string regexString, bool result)
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var regex = new Regex(regexString);
        var matcher = new RegexColumnMatcher(regex);
        Assert.Equal(result, matcher.ColumnMatches(sheet, 0));
    }

    [Fact]
    public void ColumnMatches_NullSheet_ThrowsArgumentNullException()
    {
        var regex = new Regex("Regex");
        var matcher = new RegexColumnMatcher(regex);
        Assert.Throws<ArgumentNullException>("sheet", () => matcher.ColumnMatches(null!, 0));
    }

    [Fact]
    public void ColumnMatches_SheetWithNoHeadingHasHeading_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();

        var regex = new Regex("Regex");
        var matcher = new RegexColumnMatcher(regex);
        Assert.Throws<ExcelMappingException>(() => matcher.ColumnMatches(sheet, 0));
        Assert.Null(sheet.Heading);
    }

    [Fact]
    public void ColumnMatches_SheetWithNoHeadingHasNoHeading_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;

        var regex = new Regex("Regex");
        var matcher = new RegexColumnMatcher(regex);
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

        var regex = new Regex("Regex");
        var matcher = new RegexColumnMatcher(regex);
        Assert.Throws<ArgumentOutOfRangeException>("columnIndex", () => matcher.ColumnMatches(sheet, columnIndex));
    }
}

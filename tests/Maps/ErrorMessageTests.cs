using System.Linq;

namespace ExcelMapper.Tests;

/// <summary>
/// Error messages say what's wrong in the words a spreadsheet uses: one-based rows, cell
/// addresses, the columns the heading has, and types as C# names them.
/// </summary>
public class ErrorMessageTests
{
    public class IntRow
    {
        public int Value { get; set; }
    }

    public class NullableIntRow
    {
        public int? Value { get; set; }
    }

    [Fact]
    public void ReadRows_EmptySheetWithHeading_SaysSheetIsEmpty()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet(1);

        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<IntRow>());
        Assert.Equal("Sheet \"Empty\" is empty, so it has no heading. Set HasHeading to false to read a sheet without one.", exception.Message);
    }

    [Fact]
    public void ReadHeading_HeadingRowPastEnd_SaysHowManyRows()
    {
        using var importer = Helpers.GetCsvImporter("Value\n1\n");
        var sheet = importer.ReadSheet();
        sheet.HeadingIndex = 4;

        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadHeading());
        Assert.Equal("The sheet has no row 5 to read the heading from: it has 2 rows.", exception.Message);
    }

    [Fact]
    public void ReadRows_IndexCountPastEnd_GivesOneBasedRow()
    {
        using var importer = Helpers.GetCsvImporter("Value\n1\n2\n");
        var sheet = importer.ReadSheet();

        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<IntRow>(1, 3).ToArray());
        Assert.Equal("The sheet does not have row 4.", exception.Message);
    }

    [Fact]
    public void ReadRows_MissingColumn_ListsHeadingColumns()
    {
        using var importer = Helpers.GetCsvImporter("Valeu,Other\n1,2\n");

        var exception = Assert.Throws<ExcelMappingException>(() => importer.ReadSheet().ReadRows<IntRow>().ToArray());
        Assert.Equal("Can't find column \"Value\" for member \"Value\". The heading's columns are \"Valeu\", \"Other\".", exception.Message);
        Assert.Equal(-1, exception.RowIndex);
        Assert.Equal("Value", exception.ColumnName);
    }

    [Fact]
    public void ReadRows_MissingColumnInNamedSheet_NamesSheet()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");

        var exception = Assert.Throws<ExcelMappingException>(() => importer.ReadSheet().ReadRows<IntRow>().ToArray());
        Assert.StartsWith("Can't find column \"Value\" for member \"Value\" in sheet \"Primitives\". The heading's columns are \"Int Value\", \"StringValue\",", exception.Message);
    }

    [Fact]
    public void ReadRows_MissingColumnWideHeading_ListsFirstColumns()
    {
        var heading = string.Join(",", Enumerable.Range(1, 25).Select(i => $"C{i}"));
        using var importer = Helpers.GetCsvImporter(heading + "\n" + string.Join(",", Enumerable.Range(1, 25)) + "\n");

        var exception = Assert.Throws<ExcelMappingException>(() => importer.ReadSheet().ReadRows<IntRow>().ToArray());
        Assert.EndsWith("\"C19\", \"C20\" and 5 more.", exception.Message);
    }

    [Fact]
    public void ReadRows_MissingColumnBlankHeading_SaysHeadingIsEmpty()
    {
        using var importer = Helpers.GetCsvImporter(",\n1,2\n");

        var exception = Assert.Throws<ExcelMappingException>(() => importer.ReadSheet().ReadRows<IntRow>().ToArray());
        Assert.Equal("Can't find column \"Value\" for member \"Value\". The heading is empty.", exception.Message);
    }

    [Fact]
    public void ReadRows_EmptyCell_SaysCellIsEmpty()
    {
        using var importer = Helpers.GetCsvImporter("Value,Other\n,1\n");

        var exception = Assert.Throws<ExcelMappingException>(() => importer.ReadSheet().ReadRows<IntRow>().ToArray());
        Assert.Equal("The cell is empty, but member \"Value\" of type int can't be empty in column \"Value\" at cell A2. Make it nullable (int?), or give it an empty fallback, such as [ExcelDefaultValue].", exception.Message);
    }

    [Fact]
    public void ReadRows_EmptyCellReadAsValue_NamesType()
    {
        using var importer = Helpers.GetCsvImporter("Value,Other\n,1\n");

        var exception = Assert.Throws<ExcelMappingException>(() => importer.ReadSheet().ReadRows<int>().ToArray());
        Assert.StartsWith("The cell is empty, but type int can't be empty", exception.Message);
    }

    [Fact]
    public void ReadRows_EmptyCellThrowingFallbackOnNullable_SuggestsFallback()
    {
        // A nullable member can already be empty, so making it nullable isn't suggested.
        using var importer = Helpers.GetCsvImporter("Value,Other\n,1\n");
        importer.Configuration.RegisterClassMap<NullableIntRow>(c => c.Map(p => p.Value).WithThrowingEmptyFallback());

        var exception = Assert.Throws<ExcelMappingException>(() => importer.ReadSheet().ReadRows<NullableIntRow>().ToArray());
        Assert.Equal("The cell is empty, but member \"Value\" of type int? can't be empty in column \"Value\" at cell A2. Give it an empty fallback, such as [ExcelDefaultValue].", exception.Message);
    }

    [Theory]
    [InlineData("99999999999", "Cannot assign \"99999999999\" to member \"Value\" of type int in column \"Value\" at cell A2. The value \"99999999999\" is too large or too small for int.")]
    [InlineData("2.5", "Cannot assign \"2.5\" to member \"Value\" of type int in column \"Value\" at cell A2. The value \"2.5\" isn't a whole number, so it can't be int.")]
    [InlineData("abc", "Cannot assign \"abc\" to member \"Value\" of type int in column \"Value\" at cell A2.")]
    public void ReadRows_InvalidNumber_SaysWhy(string text, string expected)
    {
        using var importer = Helpers.GetCsvImporter($"Value\n{text}\n");

        var exception = Assert.Throws<ExcelMappingException>(() => importer.ReadSheet().ReadRows<IntRow>().ToArray());
        Assert.Equal(expected, exception.Message);
    }

    public class GenericRow
    {
        public List<int?> Values { get; set; } = [];
    }

    [Fact]
    public void ValidateHeading_NestedType_UsesCSharpName()
    {
        using var importer = Helpers.GetCsvImporter("Other\n1\n");

        var exception = Assert.Throws<ExcelMappingException>(() => importer.ReadSheet().ValidateHeading<GenericRow>());
        Assert.StartsWith("The sheet is missing columns needed to map type ErrorMessageTests.GenericRow:", exception.Message);
    }

    [Fact]
    public void ReadRows_CannotAutoMap_UsesCSharpName()
    {
        using var importer = Helpers.GetCsvImporter("Value\n1\n");

        var exception = Assert.Throws<ExcelMappingException>(() => importer.ReadSheet().ReadRows<Func<int?>>().ToArray());
        Assert.Equal("Cannot auto-map type Func<int?>.", exception.Message);
    }

    [Fact]
    public void ReadRow_NoMoreRowsCsv_DoesNotNameSheet()
    {
        using var importer = Helpers.GetCsvImporter("Value\n1\n");
        var sheet = importer.ReadSheet();
        sheet.ReadRow<IntRow>();

        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IntRow>());
        Assert.Equal("There are no more rows in the sheet.", exception.Message);
    }

    [Fact]
    public void Message_ReadTwice_ReturnsSameMessage()
    {
        // The message is built when it's first read.
        using var importer = Helpers.GetCsvImporter("Value\nabc\n");
        var error = Assert.Single(Assert.Single(importer.ReadSheet().ReadRowResults<IntRow>()).Errors);
        Assert.Same(error.Message, error.Message);
        Assert.Contains(error.Message, error.ToString());
    }
}

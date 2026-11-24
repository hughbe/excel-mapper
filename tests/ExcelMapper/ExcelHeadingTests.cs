using System.Globalization;
using System.Linq;
namespace ExcelMapper.Tests;

public class ExcelHeadingTests
{
    [Fact]
    public void ColumnNames_Get_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();

        var columnNames = heading.ColumnNames;
        Assert.Equal(["Int Value", "StringValue", "Bool Value", "Enum Value", "DateValue", "ArrayValue", "MappedValue", "TrimmedValue"], columnNames);
        Assert.Same(columnNames, heading.ColumnNames);
    }

    [Fact]
    public void ColumnNames_NumbersAndDates_UseImporterCulture()
    {
        // Not the machine's culture, so a heading is named the same on every machine.
        var previous = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("de-DE");
            using var importer = XlsxBuilder.Create([1.5, 2024.0, new DateTime(2024, 3, 4), true, "Name"], [1, 2, 3, 4, "a"]);
            Assert.Equal(["1.5", "2024", "03/04/2024 00:00:00", "True", "Name"], importer.ReadSheet().ReadHeading().ColumnNames);
        }
        finally
        {
            CultureInfo.CurrentCulture = previous;
        }
    }

    [Fact]
    public void ColumnNames_NumbersWithCulture_UseCulture()
    {
        using var importer = XlsxBuilder.Create([1.5, "Name"], [1, "a"]);
        importer.Configuration.Culture = CultureInfo.GetCultureInfo("de-DE");
        Assert.Equal(["1,5", "Name"], importer.ReadSheet().ReadHeading().ColumnNames);
    }

    public class YearRow
    {
        [ExcelColumnName("2024")]
        public int Sales { get; set; }
    }

    [Fact]
    public void ReadRows_NumberHeading_MatchesColumnName()
    {
        using var importer = XlsxBuilder.Create([2023.0, 2024.0], [1, 2]);
        Assert.Equal(2, importer.ReadSheet().ReadRows<YearRow>().Single().Sales);
    }

    [Fact]
    public void GetColumnName_Invoke_Roundtrips()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();

        for (int i = 0; i < heading.ColumnNames.Count; i++)
        {
            Assert.Equal(i, heading.GetColumnIndex(heading.ColumnNames[i]));
            Assert.Equal(heading.ColumnNames[i], heading.GetColumnName(i));
        }
    }

    [Fact]
    public void GetColumnName_DuplicatedColumn_Success()
    {
        using var importer = Helpers.GetImporter("DuplicatedColumns.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();
        Assert.Equal(["MyColumn", "MyColumn"], heading.ColumnNames);
        Assert.Equal("MyColumn", heading.GetColumnName(0));
        Assert.Equal("MyColumn", heading.GetColumnName(1));
    }

    [Fact]
    public void GetColumnName_DuplicatedColumnEmpty_Success()
    {
        using var importer = Helpers.GetImporter("EmptyColumns.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();
        Assert.Equal(["", "Column2", "", " Column4 "], heading.ColumnNames);
        Assert.Equal("", heading.GetColumnName(0));
        Assert.Equal("Column2", heading.GetColumnName(1));
        Assert.Equal("", heading.GetColumnName(2));
        Assert.Equal(" Column4 ", heading.GetColumnName(3));
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(8)]
    public void GetColumnName_InvalidIndex_ThrowsArgumentOutOfRangeException(int columnIndex)
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();
        Assert.Throws<ArgumentOutOfRangeException>("columnIndex", () => heading.GetColumnName(columnIndex));
    }

    [Fact]
    public void GetColumnIndex_InvokeValidColumnName_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();

        Assert.Equal(1, heading.GetColumnIndex("StringValue"));
        Assert.Equal(1, heading.GetColumnIndex("stringvalue"));
    }

    [Fact]
    public void GetColumnIndex_DuplicatedColumn_Success()
    {
        using var importer = Helpers.GetImporter("DuplicatedColumns.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();
        Assert.Equal(["MyColumn", "MyColumn"], heading.ColumnNames);
        Assert.Equal(0, heading.GetColumnIndex("MyColumn"));
        Assert.Equal(0, heading.GetColumnIndex("mycolumn"));
    }

    [Fact]
    public void GetColumnIndex_DuplicatedColumnEmpty_Success()
    {
        using var importer = Helpers.GetImporter("EmptyColumns.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();
        Assert.Equal(["", "Column2", "", " Column4 "], heading.ColumnNames);
        Assert.Equal(0, heading.GetColumnIndex(""));
        Assert.Equal(1, heading.GetColumnIndex("Column2"));
        Assert.Equal(3, heading.GetColumnIndex(" Column4 "));
    }

    [Fact]
    public void GetColumnIndex_NullColumnName_ThrowsArgumentNullException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();

        Assert.Throws<ArgumentNullException>("columnName", () => heading.GetColumnIndex((string)null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("NoSuchColumn")]
    public void GetColumnIndex_NoSuchColumnName_ThrowsExcelMappingException(string columnName)
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => heading.GetColumnIndex(columnName));
    }

    [Fact]
    public void GetColumnName_GetColumnIndexPredicate_Roundtrips()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();

        string[] columnNames = [.. heading.ColumnNames];
        for (int i = 0; i < columnNames.Length; i++)
        {
            var scopedIndex = i;
            Assert.Equal(i, heading.GetColumnIndex(e => e == columnNames[scopedIndex]));
            Assert.Equal(columnNames[i], heading.GetColumnName(i));
        }
    }

    [Fact]
    public void GetColumnName_NullPredicate_ThrowsArgumentNullException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();

        Assert.Throws<ArgumentNullException>("predicate", () => heading.GetColumnIndex((Func<string, bool>)null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("NoSuchColumn")]
    public void GetColumnIndex_PredicateNoSuchColumnMatching_ThrowsExcelMappingException(string columnName)
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => heading.GetColumnIndex(e => e == columnName));
    }

    [Fact]
    public void TryGetColumnIndex_Invoke_Roundtrips()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();

        string[] columnNames = [.. heading.ColumnNames];
        for (int i = 0; i < columnNames.Length; i++)
        {
            Assert.True(heading.TryGetColumnIndex(columnNames[i], out int index));
            Assert.Equal(i, index);
            Assert.Equal(columnNames[i], heading.GetColumnName(i));
        }
    }

    [Fact]
    public void TryGetColumnIndex_InvokeValidColumnName_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();

        Assert.True(heading.TryGetColumnIndex("StringValue", out int index));
        Assert.Equal(1, index);
        Assert.True(heading.TryGetColumnIndex("stringvalue", out index));
        Assert.Equal(1, index);
        Assert.Equal(1, index);
    }

    [Fact]
    public void TryGetColumnIndex_DuplicatedColumn_Success()
    {
        using var importer = Helpers.GetImporter("DuplicatedColumns.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();
        Assert.Equal(["MyColumn", "MyColumn"], heading.ColumnNames);
        Assert.True(heading.TryGetColumnIndex("MyColumn", out int index));
        Assert.Equal(0, index);
        Assert.True(heading.TryGetColumnIndex("mycolumn", out index));
        Assert.Equal(0, index);
    }

    [Fact]
    public void TryGetColumnIndex_DuplicatedColumnEmpty_Success()
    {
        using var importer = Helpers.GetImporter("EmptyColumns.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();
        Assert.Equal(["", "Column2", "", " Column4 "], heading.ColumnNames);
        Assert.True(heading.TryGetColumnIndex("", out int index));
        Assert.Equal(0, index);
        Assert.True(heading.TryGetColumnIndex("Column2", out index));
        Assert.Equal(1, index);
        Assert.True(heading.TryGetColumnIndex(" Column4 ", out index));
        Assert.Equal(3, index);
    }

    [Fact]
    public void TryGetColumnIndex_NullColumnName_ThrowsArgumentNullException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();

        Assert.Throws<ArgumentNullException>("columnName", () => heading.TryGetColumnIndex((string)null!, out _));
    }

    [Theory]
    [InlineData("")]
    [InlineData("NoSuchColumn")]
    public void TryGetColumnIndex_NoSuchColumnName_ReturnsFalse(string columnName)
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();

        Assert.False(heading.TryGetColumnIndex(columnName, out int index));
        Assert.Equal(0, index);
    }

    [Fact]
    public void TryGetColumnIndex_PredicateInvoke_Roundtrips()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();

        string[] columnNames = [.. heading.ColumnNames];
        for (int i = 0; i < columnNames.Length; i++)
        {
            var scopedIndex = i;
            Assert.True(heading.TryGetColumnIndex(e => e == columnNames[scopedIndex], out int index));
            Assert.Equal(i, index);
            Assert.Equal(columnNames[i], heading.GetColumnName(i));
        }
    }

    [Fact]
    public void TryGetColumnIndex_NullPredicate_ThrowsArgumentNullException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();

        Assert.Throws<ArgumentNullException>("predicate", () => heading.TryGetColumnIndex((Func<string, bool>)null!, out _));
    }

    [Theory]
    [InlineData("")]
    [InlineData("NoSuchColumn")]
    public void TryGetColumnIndex_PredicateNoSuchColumnName_ReturnsFalse(string columnName)
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();

        Assert.False(heading.TryGetColumnIndex(e => e == columnName, out int index));
        Assert.Equal(-1, index);
    }

    [Fact]
    public void GetColumnIndex_CustomDataRangeColumnsStart_ReturnsSheetIndex()
    {
        using var importer = Helpers.GetCsvImporter("Junk,Name,Value\n1,2,3\n");
        var sheet = importer.ReadSheet();
        sheet.DataRange = new ExcelRange(Range.All, 1..);
        var heading = sheet.ReadHeading();

        Assert.Equal(["Name", "Value"], heading.ColumnNames);
        Assert.Equal(1, heading.GetColumnIndex("Name"));
        Assert.Equal(2, heading.GetColumnIndex("Value"));
        Assert.Equal("Name", heading.GetColumnName(1));
        Assert.Equal("Value", heading.GetColumnName(2));
        Assert.Throws<ArgumentOutOfRangeException>("columnIndex", () => heading.GetColumnName(0));
        Assert.Throws<ArgumentOutOfRangeException>("columnIndex", () => heading.GetColumnName(3));
        Assert.False(heading.TryGetColumnIndex("Junk", out _));
        Assert.Equal(1, heading.GetColumnIndex(_ => true));
    }

    [Fact]
    public void GetColumnIndex_PredicateNotAlphabetical_ReturnsFirstColumnInOrder()
    {
        using var importer = Helpers.GetCsvImporter("Zeta,Alpha,Beta\n1,2,3\n");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();

        Assert.Equal(0, heading.GetColumnIndex(_ => true));
        Assert.True(heading.TryGetColumnIndex(_ => true, out var index));
        Assert.Equal(0, index);
        Assert.Equal(2, heading.GetColumnIndex(n => n.StartsWith("B")));
    }

    [Fact]
    public void GetColumnIndex_PredicateDuplicatedColumn_MatchesOriginalName()
    {
        using var importer = Helpers.GetCsvImporter("Name,Name\n1,2\n");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();

        Assert.False(heading.TryGetColumnIndex(n => n == "Name_2", out _));
        Assert.Equal(0, heading.GetColumnIndex(n => n == "Name"));
    }

    [Fact]
    public void TryGetColumnIndex_DuplicatedColumnSuffix_ReturnsFalse()
    {
        // The heading doesn't invent names for repeated columns.
        using var importer = Helpers.GetCsvImporter("Name,Name\n1,2\n");
        var heading = importer.ReadSheet().ReadHeading();
        Assert.False(heading.TryGetColumnIndex("Name_2", out _));
    }

    [Fact]
    public void GetColumnIndex_ColumnNamedLikeSuffix_ReturnsColumn()
    {
        using var importer = Helpers.GetCsvImporter("Name,Name,Name_2\n1,2,3\n");
        var heading = importer.ReadSheet().ReadHeading();
        Assert.Equal(0, heading.GetColumnIndex("Name"));
        Assert.Equal(2, heading.GetColumnIndex("Name_2"));
    }

    [Fact]
    public void GetColumnIndex_MissingColumn_ListsColumnsInOrder()
    {
        using var importer = Helpers.GetCsvImporter("B,A,B\n1,2,3\n");
        var heading = importer.ReadSheet().ReadHeading();
        var exception = Assert.Throws<ExcelMappingException>(() => heading.GetColumnIndex("C"));
        Assert.Equal("The heading has no column \"C\". Its columns are \"B\", \"A\", \"B\".", exception.Message);
    }

    [Fact]
    public void GetColumnIndex_NoColumnMatchesPredicate_ThrowsExcelMappingExceptionListingColumns()
    {
        using var importer = Helpers.GetCsvImporter("A,B\n1,2\n");
        var heading = importer.ReadSheet().ReadHeading();
        var exception = Assert.Throws<ExcelMappingException>(() => heading.GetColumnIndex(n => n == "C"));
        Assert.Equal("No column's name matches the predicate. The heading's columns are \"A\", \"B\".", exception.Message);
    }
}

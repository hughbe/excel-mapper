using System.Globalization;
using System.Linq;

namespace ExcelMapper.Tests;

public class ExcelSheetTests
{
    [Fact]
    public void DataRange_GetDefault_ReturnsFalse()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        Assert.Equal(Range.All, sheet.DataRange.Rows);
        Assert.Equal(Range.All, sheet.DataRange.Columns);
    }

    public static IEnumerable<object[]> DataRange_Set_TestData()
    {
        yield return new object[] { new ExcelRange(), Range.All, Range.All };
        yield return new object[] { new ExcelRange(1, 3, 2, 4), 1..4, 2..5 };
        yield return new object[] { new ExcelRange(0, null, 0, null), 0.., 0.. };
        yield return new object[] { new ExcelRange(5, null, 1, 10), 5.., 1..11 };
        yield return new object[] { new ExcelRange(2, 6, 0, null), 2..7, 0.. };
        yield return new object[] { new ExcelRange(Range.All, Range.All), Range.All, Range.All };
        yield return new object[] { new ExcelRange(1..2, 3..4), 1..2, 3..4 };
        yield return new object[] { ExcelRange.Parse("AB2:AZ20"), 1..20, 27..52 };
    }

    [Theory]
    [MemberData(nameof(DataRange_Set_TestData))]
    public void DataRange_Set_GetReturnsExpected(ExcelRange value, Range expectedRows, Range expectedColumns)
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();

        // Set.
        sheet.DataRange = value;
        Assert.Equal(expectedRows, sheet.DataRange.Rows);
        Assert.Equal(expectedColumns, sheet.DataRange.Columns);
        Assert.Equal(expectedRows.Start.Value, sheet.HeadingIndex);

        // Set same value.
        sheet.DataRange = value;
        Assert.Equal(expectedRows, sheet.DataRange.Rows);
        Assert.Equal(expectedColumns, sheet.DataRange.Columns);
        Assert.Equal(expectedRows.Start.Value, sheet.HeadingIndex);
    }

    [Theory]
    [MemberData(nameof(DataRange_Set_TestData))]
    public void DataRange_SetHasHeadingFalse_GetReturnsExpected(ExcelRange value, Range expectedRows, Range expectedColumns)
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;

        // Set.
        sheet.DataRange = value;
        Assert.Equal(expectedRows, sheet.DataRange.Rows);
        Assert.Equal(expectedColumns, sheet.DataRange.Columns);
        Assert.Equal(expectedRows.Start.Value, sheet.HeadingIndex);

        // Set same value.
        sheet.DataRange = value;
        Assert.Equal(expectedRows, sheet.DataRange.Rows);
        Assert.Equal(expectedColumns, sheet.DataRange.Columns);
        Assert.Equal(expectedRows.Start.Value, sheet.HeadingIndex);
    }

    [Fact]
    public void DataRange_SetReadHeading_ThrowsInvalidOperationException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<InvalidOperationException>(() => sheet.DataRange = new ExcelRange());
    }

    [Fact]
    public void HasHeading_GetDefault_ReturnsTrue()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        Assert.True(sheet.HasHeading);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void HasHeading_Set_GetReturnsExpected(bool value)
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();

        // Set.
        sheet.HasHeading = value;
        Assert.Equal(value, sheet.HasHeading);
        Assert.Equal(Range.All, sheet.DataRange.Columns);

        // Set same value.
        sheet.HasHeading = value;
        Assert.Equal(value, sheet.HasHeading);

        // Set different value.
        sheet.HasHeading = !value;
        Assert.Equal(!value, sheet.HasHeading);
    }

    [Fact]
    public void HasHeading_SetWhenAlreadyRead_InvalidOperationException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<InvalidOperationException>(() => sheet.HasHeading = false);
        Assert.Throws<InvalidOperationException>(() => sheet.HasHeading = true);
    }

    [Fact]
    public void Heading_GetDefault_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        Assert.Null(sheet.Heading);
    }

    [Fact]
    public void HeadingIndex_GetDefault_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        Assert.Equal(0, sheet.HeadingIndex);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    public void HeadingIndex_Set_GetReturnsExpected(int headingIndex)
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();

        // Set.
        sheet.HeadingIndex = headingIndex;
        Assert.Equal(headingIndex, sheet.HeadingIndex);
        Assert.Equal(headingIndex.., sheet.DataRange.Rows);
        Assert.Equal(Range.All, sheet.DataRange.Columns);

        // Set same value.
        sheet.HeadingIndex = headingIndex;
        Assert.Equal(headingIndex, sheet.HeadingIndex);
        Assert.Equal(headingIndex.., sheet.DataRange.Rows);
    }

    [Fact]
    public void HeadingIndex_SetNegative_ThrowsArgumentOutOfRangeException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        Assert.Throws<ArgumentOutOfRangeException>("value", () => sheet.HeadingIndex = -1);

        sheet.HasHeading = false;
        Assert.Throws<ArgumentOutOfRangeException>("value", () => sheet.HeadingIndex = -1);

        sheet.HasHeading = true;
        sheet.ReadHeading();
        Assert.Throws<ArgumentOutOfRangeException>("value", () => sheet.HeadingIndex = -1);
    }

    [Fact]
    public void HeadingIndex_SetAfterHeadingSet_ThrowsInvalidOperationException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<InvalidOperationException>(() => sheet.HeadingIndex = 0);
    }

    [Fact]
    public void HeadingIndex_SetWhenHasHeadingFalse_ThrowsInvalidOperationException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;

        Assert.Throws<InvalidOperationException>(() => sheet.HeadingIndex = 0);
    }

    [Fact]
    public void NumberOfColumns_Get_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        Assert.Equal(1, sheet.NumberOfColumns);
    }

    [Fact]
    public void Visibility_Get_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("HiddenSheets.xlsx");
        var sheet = importer.ReadSheet();
        Assert.Equal("VisibleSheet", sheet.Name);
        Assert.Equal(ExcelSheetVisibility.Visible, sheet.Visibility);

        sheet = importer.ReadSheet();
        Assert.Equal("VeryHiddenSheet", sheet.Name);
        Assert.Equal(ExcelSheetVisibility.VeryHidden, sheet.Visibility);

        sheet = importer.ReadSheet();
        Assert.Equal("HiddenSheet", sheet.Name);
        Assert.Equal(ExcelSheetVisibility.Hidden, sheet.Visibility);
    }

    [Fact]
    public void ReadHeading_HasHeading_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();
        Assert.Same(heading, sheet.Heading);
        Assert.Equal(["Int Value", "StringValue", "Bool Value", "Enum Value", "DateValue", "ArrayValue", "MappedValue", "TrimmedValue"], heading.ColumnNames);
        Assert.Equal(0, sheet.CurrentRowIndex);
    }

    [Fact]
    public void ReadHeading_DuplicatedColumnName_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("DuplicatedColumns.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();
        Assert.Same(heading, sheet.Heading);
        Assert.Equal(["MyColumn", "MyColumn"], heading.ColumnNames);
        Assert.Equal(0, sheet.CurrentRowIndex);
    }

    [Fact]
    public void ReadHeading_EmptyColumnName_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("EmptyColumns.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.ReadHeading();
        Assert.Same(heading, sheet.Heading);
        Assert.Equal(["", "Column2", "", " Column4 "], heading.ColumnNames);
        Assert.Equal(0, sheet.CurrentRowIndex);
    }

    [Fact]
    public void ReadHeading_NonZeroHeadingIndex_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Non Zero Header Index.xlsx");
        var sheet = importer.ReadSheet();
        sheet.HeadingIndex = 3;

        var heading = sheet.ReadHeading();
        Assert.Same(heading, sheet.Heading);
        Assert.Equal(["Value"], heading.ColumnNames);
        Assert.Equal(3, sheet.CurrentRowIndex);
    }

    [Fact]
    public void ReadHeading_CustomDataRangeRowsStart_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Non Zero Header Index.xlsx");
        var sheet = importer.ReadSheet();
        sheet.DataRange = new ExcelRange(3.., Range.All);

        var heading = sheet.ReadHeading();
        Assert.Same(heading, sheet.Heading);
        Assert.Equal(["Value"], heading.ColumnNames);
        Assert.Equal(3, sheet.CurrentRowIndex);
    }

    [Fact]
    public void ReadHeading_CustomDataRangeColumnsStart_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        sheet.DataRange = new ExcelRange(Range.All, 1..);

        var heading = sheet.ReadHeading();
        Assert.Same(heading, sheet.Heading);
        Assert.Equal(["StringValue", "Bool Value", "Enum Value", "DateValue", "ArrayValue", "MappedValue", "TrimmedValue"], heading.ColumnNames);
        Assert.Equal(0, sheet.CurrentRowIndex);
    }

    [Fact]
    public void ReadHeading_CustomDataRangeColumnsEnd_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        sheet.DataRange = new ExcelRange(Range.All, ..^1);

        var heading = sheet.ReadHeading();
        Assert.Same(heading, sheet.Heading);
        Assert.Equal(["Int Value", "StringValue", "Bool Value", "Enum Value", "DateValue", "ArrayValue", "MappedValue"], heading.ColumnNames);
        Assert.Equal(0, sheet.CurrentRowIndex);
    }

    [Fact]
    public void ReadHeading_CustomDataRangeColumnsStartEnd_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        sheet.DataRange = new ExcelRange(Range.All, 1..4);

        var heading = sheet.ReadHeading();
        Assert.Same(heading, sheet.Heading);
        Assert.Equal(["StringValue", "Bool Value", "Enum Value"], heading.ColumnNames);
        Assert.Equal(0, sheet.CurrentRowIndex);
    }

    [Fact]
    public void ReadHeading_CustomDataRangeColumnsLargeStart_ThrowsArgumentOutOfRangeException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        sheet.DataRange = ExcelRange.Parse("I:J");

        Assert.Throws<ArgumentOutOfRangeException>("length", () => sheet.ReadHeading());
    }

    [Fact]
    public void ReadHeading_CustomDataRangeColumnsLargeEnd_ThrowsArgumentOutOfRangeException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        sheet.DataRange = ExcelRange.Parse("A:J");

        Assert.Throws<ArgumentOutOfRangeException>("length", () => sheet.ReadHeading());
    }

    [Fact]
    public void ReadHeading_AlreadyReadHeading_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadHeading());
    }

    [Fact]
    public void ReadHeading_DoesNotHaveHasHeading_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;

        Assert.Throws<ExcelMappingException>(() => sheet.ReadHeading());
    }

    [Fact]
    public void ReadHeading_NoRows_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.ReadSheet();

        ExcelSheet emptySheet = importer.ReadSheet();
        Assert.Throws<ExcelMappingException>(() => emptySheet.ReadHeading());
    }

    [Fact]
    public void ReadHeading_TooLargeHeadingIndex_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Non Zero Header Index.xlsx");
        var sheet = importer.ReadSheet();
        sheet.HeadingIndex = 8;

        Assert.Throws<ExcelMappingException>(() => sheet.ReadHeading());
    }

    [Fact]
    public void ReadHeading_Closed_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        importer.Dispose();

        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadHeading());
        Assert.Equal("The underlying reader is closed.", ex.Message);
    }

    [Fact]
    public void ReadHeading_TooManyColumns_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.MaxColumnsPerSheet = 5; // Primitives.xlsx has 8 columns
        var sheet = importer.ReadSheet();

        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadHeading());
        Assert.Contains("exceeds the maximum", ex.Message);
        Assert.Contains("8 columns", ex.Message);
        Assert.Contains("(5)", ex.Message);
    }

    [Fact]
    public void ReadHeading_ExactlyMaxColumns_Success()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.MaxColumnsPerSheet = 8; // Primitives.xlsx has exactly 8 columns
        var sheet = importer.ReadSheet();

        var heading = sheet.ReadHeading();
        Assert.NotNull(heading);
        Assert.Equal(8, heading.ColumnNames.Count);
    }

    [Fact]
    public void ReadHeading_DefaultMaxColumns_AllowsLargeFiles()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();

        var heading = sheet.ReadHeading();
        Assert.NotNull(heading);
        Assert.Equal(8, heading.ColumnNames.Count);
    }

    [Fact]
    public void ReadRows_NotReadHeading_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();

        var rows = sheet.ReadRows<StringValue>();
        Assert.Equal(new string?[] { "value", "  value  ", null, "value" }, rows.Select(p => p.Value).ToArray());
        Assert.Equal(4, sheet.CurrentRowIndex);

        Assert.NotNull(sheet.Heading);
        Assert.True(sheet.HasHeading);
    }

    [Fact]
    public void ReadRows_HasHeadingFalse_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<StringValueClassMapColumnIndex>();

        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;

        var rows = sheet.ReadRows<StringValue>();
        Assert.Equal(new string?[] { "Value", "value", "  value  ", null, "value" }, rows.Select(p => p.Value).ToArray());
        Assert.Equal(4, sheet.CurrentRowIndex);

        Assert.Null(sheet.Heading);
        Assert.False(sheet.HasHeading);
    }

    [Fact]
    public void ReadRows_ReadHeading_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var rows = sheet.ReadRows<StringValue>();
        Assert.Equal(new string?[] { "value", "  value  ", null, "value"  }, rows.Select(p => p.Value).ToArray());
        Assert.Equal(4, sheet.CurrentRowIndex);

        Assert.NotNull(sheet.Heading);
        Assert.True(sheet.HasHeading);
    }

    [Fact]
    public void ReadRows_ReadHeadingNonZeroHeadingIndex_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Non Zero Header Index.xlsx");
        var sheet = importer.ReadSheet();
        sheet.HeadingIndex = 3;
        sheet.ReadHeading();

        var rows = sheet.ReadRows<StringValue>();
        Assert.Equal(new string?[] { "value", "  value  ", null, "value" }, rows.Select(p => p.Value).ToArray());
        Assert.Equal(7, sheet.CurrentRowIndex);

        Assert.NotNull(sheet.Heading);
        Assert.True(sheet.HasHeading);

        // No more rows to read.
        Assert.Empty(sheet.ReadRows<StringValue>());
    }

    [Fact]
    public void ReadRows_CustomDataRangeRows_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Non Zero Header Index.xlsx");
        var sheet = importer.ReadSheet();
        sheet.DataRange = ExcelRange.Parse("4:7");
        sheet.ReadHeading();

        var rows = sheet.ReadRows<StringValue>();
        Assert.Equal(new string?[] { "value", "  value  ", null }, rows.Select(p => p.Value).ToArray());
        Assert.Equal(6, sheet.CurrentRowIndex);

        Assert.NotNull(sheet.Heading);
        Assert.True(sheet.HasHeading);

        // No more rows to read.
        Assert.Empty(sheet.ReadRows<StringValue>());
        Assert.Equal(6, sheet.CurrentRowIndex);
    }

    [Fact]
    public void ReadRows_CustomDataRangeRowsOutOfRange_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Non Zero Header Index.xlsx");
        var sheet = importer.ReadSheet();
        sheet.DataRange = ExcelRange.Parse("4:9");
        sheet.ReadHeading();

        // The sheet ends before the range: the last row isn't repeated.
        var rows = sheet.ReadRows<StringValue>();
        Assert.Equal(new string?[] { "value", "  value  ", null, "value" }, rows.Select(p => p.Value).ToArray());
        Assert.Equal(7, sheet.CurrentRowIndex);

        Assert.NotNull(sheet.Heading);
        Assert.True(sheet.HasHeading);

        // No more rows to read.
        Assert.Empty(sheet.ReadRows<StringValue>());
        Assert.Equal(7, sheet.CurrentRowIndex);
    }

    [Fact]
    public void ReadRows_CustomDataRangeColumns_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Primitives_Range.xlsx");
        var sheet = importer.ReadSheet();
        sheet.DataRange = ExcelRange.Parse("B:D");
        sheet.ReadHeading();

        var rows = sheet.ReadRows<string[]>().ToArray();
        Assert.Equal(new string?[] { "a", "True", "Government" }, rows[0]);
        Assert.Equal(4, sheet.CurrentRowIndex);

        Assert.NotNull(sheet.Heading);
        Assert.True(sheet.HasHeading);

        // No more rows to read.
        Assert.Empty(sheet.ReadRows<StringValue>());
        Assert.Equal(4, sheet.CurrentRowIndex);
    }

    [Fact]
    public void ReadRows_CustomDataRangeColumnsLargeStart_ThrowsArgumentOutOfRangeException()
    {
        using var importer = Helpers.GetImporter("Primitives_Range.xlsx");
        var sheet = importer.ReadSheet();
        sheet.DataRange = ExcelRange.Parse("I:J");

        Assert.Throws<ArgumentOutOfRangeException>("length", () => sheet.ReadRows<StringValue>());
        Assert.Equal(0, sheet.CurrentRowIndex);

        // Read again.
        Assert.Throws<ArgumentOutOfRangeException>("length", () => sheet.ReadRows<StringValue>());
        Assert.Equal(0, sheet.CurrentRowIndex);
    }

    [Fact]
    public void ReadRows_CustomDataRangeColumnsLargeEnd_ThrowsArgumentOutOfRangeException()
    {
        using var importer = Helpers.GetImporter("Primitives_Range.xlsx");
        var sheet = importer.ReadSheet();
        sheet.DataRange = ExcelRange.Parse("A:J");

        Assert.Throws<ArgumentOutOfRangeException>("length", () => sheet.ReadRows<StringValue>());
        Assert.Equal(0, sheet.CurrentRowIndex);

        // Read again.
        Assert.Throws<ArgumentOutOfRangeException>("length", () => sheet.ReadRows<StringValue>());
        Assert.Equal(0, sheet.CurrentRowIndex);
    }

    [Fact]
    public void ReadRows_CustomDataRangeColumnNameOutOfRange_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives_Range.xlsx");
        importer.Configuration.RegisterClassMap<StringValue>(c =>
        {
            c.Map(p => p.Value)
                .WithColumnName("Int Value");
        });

        var sheet = importer.ReadSheet();
        sheet.DataRange = ExcelRange.Parse("B1:D4");
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<StringValue>().ToArray());
        Assert.Equal(1, sheet.CurrentRowIndex);

        // Read again.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<StringValue>().ToArray());
        Assert.Equal(2, sheet.CurrentRowIndex);

        // Read again.
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<StringValue>().ToArray());
        Assert.Equal(3, sheet.CurrentRowIndex);

        // No more rows to read.
        Assert.Empty(sheet.ReadRows<StringValue>());
        Assert.Equal(3, sheet.CurrentRowIndex);

        // Read again.
        Assert.Empty(sheet.ReadRows<StringValue>());
        Assert.Equal(3, sheet.CurrentRowIndex);
    }

    [Fact]
    public void ReadRows_AllReadHasHeadingTrue_ReturnsEmpty()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();

        var rows1 = sheet.ReadRows<StringValue>();
        Assert.Equal(new string?[] { "value", "  value  ", null, "value" }, rows1.Select(p => p.Value).ToArray());
        Assert.Equal(4, sheet.CurrentRowIndex);

        Assert.NotNull(sheet.Heading);
        Assert.True(sheet.HasHeading);

        // No more rows to read.
        Assert.Empty(sheet.ReadRows<StringValue>().ToArray());
        Assert.Equal(4, sheet.CurrentRowIndex);
    }

    [Fact]
    public void ReadRows_AllReadingHasHeadingFalse_ReturnsEmpty()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<StringValueClassMapColumnIndex>();

        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;

        var rows1 = sheet.ReadRows<StringValue>();
        Assert.Equal(new string?[] { "Value", "value", "  value  ", null, "value" }, rows1.Select(p => p.Value).ToArray());
        Assert.Equal(4, sheet.CurrentRowIndex);

        Assert.Null(sheet.Heading);
        Assert.False(sheet.HasHeading);

        // No more rows to read.
        Assert.Empty(sheet.ReadRows<StringValue>().Select(p => p.Value).ToArray());
        Assert.Equal(4, sheet.CurrentRowIndex);
    }

    [Fact]
    public void ReadRows_EmptySheetNoHeading_ReturnsEmpty()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<StringValueClassMapColumnIndex>();

        importer.ReadSheet();

        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;

        var rows = sheet.ReadRows<StringValue>();
        Assert.Empty(rows.ToArray());
    }

    [Fact]
    public void ReadRows_EmptySheetHeading_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<StringValueClassMapColumnIndex>();

        importer.ReadSheet();

        var sheet = importer.ReadSheet();
        sheet.HasHeading = true;

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<StringValue>());
    }

    [Fact]
    public void ReadRows_Closed_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        importer.Dispose();

        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<StringValue>());
        Assert.Equal("The underlying reader is closed.", ex.Message);
    }

    public static IEnumerable<object[]> ReadRows_IndexCount_TestData()
    {
        yield return new object[] { 1, 4, new string?[] { "value", "  value  ", null, "value" } };
        yield return new object[] { 1, 3, new string?[] { "value", "  value  ", null } };
        yield return new object[] { 1, 2, new string?[] { "value", "  value  " } };
        yield return new object[] { 2, 1, new string?[] { "  value  " } };
        yield return new object[] { 1, 0, Array.Empty<string?>() };
    }

    [Theory]
    [MemberData(nameof(ReadRows_IndexCount_TestData))]
    public void ReadRows_IndexCount_ReturnsExpected(int startIndex, int count, string?[] expectedValues)
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var rows = sheet.ReadRows<StringValue>(startIndex, count);
        Assert.Equal(expectedValues, rows.Select(p => p.Value).ToArray());
        Assert.Equal(count == 0 ? 0 : startIndex + count - 1, sheet.CurrentRowIndex);

        Assert.NotNull(sheet.Heading);
        Assert.True(sheet.HasHeading);
    }

    [Theory]
    [MemberData(nameof(ReadRows_IndexCount_TestData))]
    public void ReadRows_IndexCountNotReadHeading_ReturnsExpected(int startIndex, int count, string?[] expectedValues)
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();

        var rows = sheet.ReadRows<StringValue>(startIndex, count);
        Assert.Equal(expectedValues, rows.Select(p => p.Value).ToArray());
        Assert.Equal(count == 0 ? 0 : startIndex + count - 1, sheet.CurrentRowIndex);

        Assert.NotNull(sheet.Heading);
        Assert.True(sheet.HasHeading);
    }

    [Theory]
    [MemberData(nameof(ReadRows_IndexCount_TestData))]
    public void ReadRows_IndexCountReadHeading_ReturnsExpected(int startIndex, int count, string?[] expectedValues)
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();

        var heading = sheet.ReadHeading();

        var rows = sheet.ReadRows<StringValue>(startIndex, count);
        Assert.Equal(expectedValues, rows.Select(p => p.Value).ToArray());
        Assert.Equal(count == 0 ? 0 : startIndex + count - 1, sheet.CurrentRowIndex);

        Assert.NotNull(sheet.Heading);
        Assert.Same(heading, sheet.Heading);
        Assert.True(sheet.HasHeading);
    }

    public static IEnumerable<object[]> ReadRows_IndexCountNoHeading_TestData()
    {
        yield return new object[] { 0, 5, new string?[] { "Value", "value", "  value  ", null, "value" } };
        yield return new object[] { 0, 4, new string?[] { "Value", "value", "  value  ", null } };
        yield return new object[] { 1, 4, new string?[] { "value", "  value  ", null, "value" } };
        yield return new object[] { 1, 2, new string?[] { "value", "  value  " } };
        yield return new object[] { 1, 0, Array.Empty<string?>() };
    }

    [Theory]
    [MemberData(nameof(ReadRows_IndexCountNoHeading_TestData))]
    public void ReadRows_IndexCountNoHeading_ReturnsExpected(int startIndex, int count, string?[] expectedValues)
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<StringValueClassMapColumnIndex>();
        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;

        var rows = sheet.ReadRows<StringValue>(startIndex, count).ToList();
        Assert.Equal(expectedValues, rows.Select(p => p.Value).ToArray());
        Assert.Equal(count == 0 ? 0 : startIndex + count - 1, sheet.CurrentRowIndex);

        Assert.Null(sheet.Heading);
        Assert.False(sheet.HasHeading);
    }

    [Fact]
    public void ReadRows_IntIntClosed_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        importer.Dispose();

        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<StringValue>(0, 1));
        Assert.Equal("The underlying reader is closed.", ex.Message);
    }

    [Fact]
    public void ReadRows_IntIntEmptySheetNoHeadingZero_ReturnsEmpty()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<StringValueClassMapColumnIndex>();

        importer.ReadSheet();

        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;

        var rows = sheet.ReadRows<StringValue>(0, 0);
        Assert.Empty(rows);
    }

    [Fact]
    public void ReadRows_IntIntEmptySheetNoHeadingNonZeroCount_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<StringValueClassMapColumnIndex>();

        importer.ReadSheet();

        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;

        var rows = sheet.ReadRows<StringValue>(0, 1);
        Assert.Throws<ExcelMappingException>(() => rows.ToArray());
    }

    [Fact]
    public void ReadRows_IntIntEmptySheetHeading_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.RegisterClassMap<StringValueClassMapColumnIndex>();

        importer.ReadSheet();

        var sheet = importer.ReadSheet();
        sheet.HasHeading = true;

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<StringValue>(1, 1));
    }

    [Fact]
    public void ReadRows_BlankLinesNotSkipped_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("BlankLines.xlsx");
        var sheet = importer.ReadSheet();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<BlankLinesClass>().ToArray());
        Assert.Equal(1, sheet.CurrentRowIndex);
    }

    [Fact]
    public void ReadRows_BlankLinesSkipped_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("BlankLines.xlsx");
        importer.Configuration.SkipBlankLines = true;
        var sheet = importer.ReadSheet();

        BlankLinesClass[] rows = [.. sheet.ReadRows<BlankLinesClass>()];
        Assert.Equal(4, rows.Length);
        Assert.Equal("A", rows[0].StringValue);
        Assert.Equal(1, rows[0].IntValue);
        Assert.Equal("B", rows[1].StringValue);
        Assert.Equal(2, rows[1].IntValue);
        Assert.Null(rows[2].StringValue);
        Assert.Equal(3, rows[2].IntValue);
        Assert.Equal("C", rows[3].StringValue);
        Assert.Equal(0, rows[3].IntValue);
        // The sheet's formatted blank rows after the last value aren't rows.
        Assert.Equal(9, sheet.CurrentRowIndex);
    }

    [Fact]
    public void ReadRow_BlankLinesNotSkipped_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("BlankLines.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<BlankLinesClass>());
        Assert.Equal(1, sheet.CurrentRowIndex);

        var row1 = sheet.ReadRow<BlankLinesClass>();
        Assert.Equal("A", row1.StringValue);
        Assert.Equal(1, row1.IntValue);
        Assert.Equal(2, sheet.CurrentRowIndex);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<BlankLinesClass>());
        Assert.Equal(3, sheet.CurrentRowIndex);

        var row2 = sheet.ReadRow<BlankLinesClass>();
        Assert.Equal("B", row2.StringValue);
        Assert.Equal(2, row2.IntValue);
        Assert.Equal(4, sheet.CurrentRowIndex);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<BlankLinesClass>());
        Assert.Equal(5, sheet.CurrentRowIndex);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<BlankLinesClass>());
        Assert.Equal(6, sheet.CurrentRowIndex);

        var row3 = sheet.ReadRow<BlankLinesClass>();
        Assert.Null(row3.StringValue);
        Assert.Equal(3, row3.IntValue);
        Assert.Equal(7, sheet.CurrentRowIndex);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<BlankLinesClass>());
        Assert.Equal(8, sheet.CurrentRowIndex);

        var row4 = sheet.ReadRow<BlankLinesClass>();
        Assert.Equal("C", row4.StringValue);
        Assert.Equal(0, row4.IntValue);
        Assert.Equal(9, sheet.CurrentRowIndex);

        // The sheet's blank rows after row 10 aren't rows, so there are no more.
        var exception = Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<BlankLinesClass>());
        Assert.Contains("no more rows", exception.Message);
        Assert.Equal(9, sheet.CurrentRowIndex);
    }

    [Fact]
    public void ReadRow_BlankLinesSkipped_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("BlankLines.xlsx");
        importer.Configuration.SkipBlankLines = true;
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<BlankLinesClass>();
        Assert.Equal("A", row1.StringValue);
        Assert.Equal(1, row1.IntValue);
        Assert.Equal(2, sheet.CurrentRowIndex);

        var row2 = sheet.ReadRow<BlankLinesClass>();
        Assert.Equal("B", row2.StringValue);
        Assert.Equal(2, row2.IntValue);
        Assert.Equal(4, sheet.CurrentRowIndex);

        var row3 = sheet.ReadRow<BlankLinesClass>();
        Assert.Null(row3.StringValue);
        Assert.Equal(3, row3.IntValue);
        Assert.Equal(7, sheet.CurrentRowIndex);

        var row4 = sheet.ReadRow<BlankLinesClass>();
        Assert.Equal("C", row4.StringValue);
        Assert.Equal(0, row4.IntValue);
        Assert.Equal(9, sheet.CurrentRowIndex);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<BlankLinesClass>());
        // The sheet's formatted blank rows after the last value aren't rows.
        Assert.Equal(9, sheet.CurrentRowIndex);
    }

    [Fact]
    public void ReadRow_BlankLinesEmptySkipped_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("BlankLines_Empty.xlsx");
        importer.Configuration.SkipBlankLines = true;
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        var row1 = sheet.ReadRow<BlankLinesClass>();
        Assert.Equal("A", row1.StringValue);
        Assert.Equal(1, row1.IntValue);
        Assert.Equal(2, sheet.CurrentRowIndex);

        var row2 = sheet.ReadRow<BlankLinesClass>();
        Assert.Equal("B", row2.StringValue);
        Assert.Equal(2, row2.IntValue);
        Assert.Equal(4, sheet.CurrentRowIndex);

        var row3 = sheet.ReadRow<BlankLinesClass>();
        Assert.Null(row3.StringValue);
        Assert.Equal(3, row3.IntValue);
        Assert.Equal(7, sheet.CurrentRowIndex);

        var row4 = sheet.ReadRow<BlankLinesClass>();
        Assert.Equal("C", row4.StringValue);
        Assert.Equal(0, row4.IntValue);
        Assert.Equal(9, sheet.CurrentRowIndex);

        var row5 = sheet.ReadRow<BlankLinesClass>();
        Assert.Equal(new DateTime(2025, 10, 25, 0, 0, 0).ToString(CultureInfo.InvariantCulture), row5.StringValue);
        Assert.Equal(4, row5.IntValue);
        Assert.Equal(999, sheet.CurrentRowIndex);

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<BlankLinesClass>());
        Assert.Equal(999, sheet.CurrentRowIndex);
    }

    [Fact]
    public void ReadRows_NegativeStartIndex_ThrowsArgumentOutOfRangeException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        Assert.Throws<ArgumentOutOfRangeException>("startIndex", () => sheet.ReadRows<StringValue>(-1, 0));
    }

    [Theory]
    [InlineData(0, 0)]
    [InlineData(1, 0)]
    [InlineData(1, 1)]
    public void ReadRows_StartIndexLargerThanHeadingIndex_ThrowsArgumentOutOfRangeException(int headingIndex, int startIndex)
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.HeadingIndex = headingIndex;
        Assert.Throws<ArgumentOutOfRangeException>("startIndex", () => sheet.ReadRows<StringValue>(startIndex, 0));
    }

    [Fact]
    public void ReadRows_NegativeCount_ThrowsArgumentOutOfRangeException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        Assert.Throws<ArgumentOutOfRangeException>("count", () => sheet.ReadRows<StringValue>(1, -1));
    }

    [Fact]
    public void ReadRows_LargeCount_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<StringValue>(1, 1000).ToArray());
        Assert.Equal(4, sheet.CurrentRowIndex);
    }

    [Fact]
    public void ReadRows_LargeCountOffsetFromHeading_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<StringValue>(1000, 1000).ToArray());
        Assert.Equal(4, sheet.CurrentRowIndex);
    }

    [Fact]
    public void ReadRow_HasHeadingFalse_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<StringValueClassMapColumnIndex>();

        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;

        StringValue value = sheet.ReadRow<StringValue>();
        Assert.Equal("Value", value.Value);
        Assert.Equal(0, sheet.CurrentRowIndex);

        Assert.Null(sheet.Heading);
        Assert.False(sheet.HasHeading);
    }

    [Fact]
    public void ReadRow_HasHeadingFalseAutomapped_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<StringValue>());
    }

    [Fact]
    public void ReadRow_HasHeadingFalseColumnNameMapping_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<StringValueClassMapColumnName>();

        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<StringValue>());
    }

    [Fact]
    public void ReadRow_HasHeadingFalseColumnNamesMapping_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<StringValuesClassMapColumnNames>();

        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<StringValues>());
    }

    private class StringValueClassMapColumnIndex : ExcelClassMap<StringValue>
    {
        public StringValueClassMapColumnIndex()
        {
            Map(value => value.Value)
                .WithColumnIndex(0);
        }
    }

    private class StringValueClassMapColumnName : ExcelClassMap<StringValue>
    {
        public StringValueClassMapColumnName()
        {
            Map(value => value.Value)
                .WithColumnName("Value");
        }
    }

    private class StringValuesClassMapColumnNames : ExcelClassMap<StringValues>
    {
        public StringValuesClassMapColumnNames()
        {
            Map(value => value.Value)
                .WithColumnNames("Value");
        }
    }

    private class StringValue
    {
        public string? Value { get; set; }
    }

    private class StringValues
    {
        public string[] Value { get; set; } = default!;
    }

    [Fact]
    public void ReadRow_CantMapType_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Helpers.IListInterface>());
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IDisposable>());
    }

    [Fact]
    public void ReadRow_NoMoreRows_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        Assert.NotEmpty(sheet.ReadRows<object>().ToArray());

        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<object>());
        Assert.Equal("There are no more rows in sheet \"Primitives\".", ex.Message);
    }

    [Fact]
    public void TryReadRow_NoMoreRows_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        var sheet = importer.ReadSheet();
        Assert.NotEmpty(sheet.ReadRows<object>().ToArray());

        Assert.False(sheet.TryReadRow(out object? row));
        Assert.Null(row);
    }

    [Fact]
    public void ReadRow_Closed_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        importer.Dispose();

        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<StringValue>());
        Assert.Equal("The underlying reader is closed.", ex.Message);
    }

    [Fact]
    public void TryReadRow_Closed_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        importer.Dispose();

        var ex = Assert.Throws<ExcelMappingException>(() => sheet.TryReadRow<StringValue>(out var row));
        Assert.Equal("The underlying reader is closed.", ex.Message);
    }

    [Fact]
    public void Region_Invoke_ReturnsWholeWorksheet()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        var region = sheet.Region();
        Assert.Same(sheet, region.Sheet);
        Assert.Equal(default, region.Range);
        Assert.NotSame(region, sheet.Region());
    }

    public static IEnumerable<object[]> Region_ExcelRange_TestData()
    {
        yield return new object[] { ExcelRange.Parse("B3:F20") };
        yield return new object[] { new ExcelRange(1..5, 0..2) };
        yield return new object[] { new ExcelRange(..^1, Range.All) };
    }

    [Theory]
    [MemberData(nameof(Region_ExcelRange_TestData))]
    public void Region_ExcelRange_ReturnsExpected(ExcelRange range)
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        var region = sheet.Region(range);
        Assert.Same(sheet, region.Sheet);
        Assert.Equal(range, region.Range);
    }

    [Theory]
    [InlineData("B3:F20")]
    [InlineData("$B$3:$F$20")]
    [InlineData("B:F")]
    [InlineData("3:20")]
    public void Region_String_ReturnsExpected(string address)
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        var region = sheet.Region(address);
        Assert.Same(sheet, region.Sheet);
        Assert.Equal(ExcelRange.Parse(address), region.Range);
    }

    [Fact]
    public void Region_NullAddress_ThrowsArgumentNullException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        Assert.Throws<ArgumentNullException>("address", () => sheet.Region((string)null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("A2:B")]
    [InlineData("invalid")]
    public void Region_InvalidAddress_ThrowsArgumentException(string address)
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        Assert.Throws<ArgumentException>("address", () => sheet.Region(address));
    }

    [Fact]
    public void Region_AfterReadingHeading_DoesNotUseSheetHeading()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();
        var values = sheet.Region().ReadRows<StringValue>().Select(v => v.Value).ToArray();
        Assert.Equal(sheet.ReadRows<StringValue>().Select(v => v.Value).ToArray(), values);
    }

    [Fact]
    public void FindHeading_Invoke_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        var heading = sheet.FindHeading<StringValue>();
        Assert.Same(heading, sheet.Heading);
        Assert.Equal(["Value"], heading.ColumnNames);
        Assert.Equal(0, sheet.HeadingIndex);
        Assert.True(sheet.HasHeading);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    public void FindHeading_InvalidMaxRows_ThrowsArgumentOutOfRangeException(int maxRows)
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        Assert.Throws<ArgumentOutOfRangeException>("maxRows", () => sheet.FindHeading<StringValue>(maxRows));
        Assert.Throws<ArgumentOutOfRangeException>("maxRows", () => sheet.TryFindHeading<StringValue>(maxRows, out _));
    }

    [Fact]
    public void FindHeading_AlreadyReadHeading_ThrowsInvalidOperationException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        sheet.ReadHeading();
        Assert.Throws<InvalidOperationException>(() => sheet.FindHeading<StringValue>());
        Assert.Throws<InvalidOperationException>(() => sheet.TryFindHeading<StringValue>(out _));
    }

    [Fact]
    public void FindHeading_Closed_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        importer.Dispose();

        var ex = Assert.Throws<ExcelMappingException>(() => sheet.FindHeading<StringValue>());
        Assert.Equal("The underlying reader is closed.", ex.Message);
        ex = Assert.Throws<ExcelMappingException>(() => sheet.TryFindHeading<StringValue>(out _));
        Assert.Equal("The underlying reader is closed.", ex.Message);
    }

    private class BlankLinesClass
    {
        public string? StringValue { get; set; }
        public int IntValue { get; set; }
    }
    class Dictionary
    {
        public Dictionary<string, string> RawRow { get; set; } = default!;
    }

    [Fact]
    public void ReadDictionary_FormattedCellOutsideRange()
    {
        // One of the cells header row outside of the data range has a bold formatting applied to it
        // causing the ManyToOne Cell reader to add a header with an empty value
        using var importer = Helpers.GetImporter("DictionaryMappingIssue.xlsx");
        var map = new ExcelClassMap<Dictionary>();

        map.Map(t => t.RawRow);
        importer.Configuration.RegisterClassMap(map);
        var sheet = importer.ReadSheet();

        var rows = sheet.ReadRows<Dictionary>().ToArray();

        Assert.Equal(3, rows.Count());
        Assert.Equal(4, rows[0].RawRow.Count());
    }

    [Fact]
    public void ReadRows_ExceedsMaxColumns_ThrowsExcelMappingException()
    {
        using var importer = Helpers.GetImporter("Primitives.xlsx");
        importer.Configuration.MaxColumnsPerSheet = 3; // Primitives.xlsx has 8 columns
        var sheet = importer.ReadSheet();

        // Exception should be thrown when ReadHeading is called (during ReadRows if HasHeading is true)
        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<StringValue>().ToList());
        Assert.Contains("exceeds the maximum", ex.Message);
    }

    [Fact]
    public void ReadRows_WithinMaxColumns_Success()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.MaxColumnsPerSheet = 100; // Strings.xlsx has 1 column, well within limit
        var sheet = importer.ReadSheet();

        var rows = sheet.ReadRows<StringValue>().ToList();
        Assert.NotEmpty(rows);
    }

    [Fact]
    public void ReadRows_BlankLinesSkippedDataRangeEndsInBlankRows_DoesNotReadPastRange()
    {
        using var importer = Helpers.GetImporter("BlankLines.xlsx");
        importer.Configuration.SkipBlankLines = true;
        var sheet = importer.ReadSheet();
        // Rows 0-5: heading, blank, "A", blank, "B", blank. Rows 6+ must not be read.
        sheet.DataRange = new ExcelRange(0..6, Range.All);

        BlankLinesClass[] rows = [.. sheet.ReadRows<BlankLinesClass>()];
        Assert.Equal(2, rows.Length);
        Assert.Equal("A", rows[0].StringValue);
        Assert.Equal("B", rows[1].StringValue);
        Assert.True(sheet.CurrentRowIndex <= 5);
    }

    [Fact]
    public void ReadRows_IndexCountCalledTwice_RewindsToStartIndex()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();

        Assert.Equal(new string?[] { "value", "  value  " }, sheet.ReadRows<StringValue>(1, 2).Select(p => p.Value).ToArray());
        Assert.Equal(new string?[] { "value", "  value  " }, sheet.ReadRows<StringValue>(1, 2).Select(p => p.Value).ToArray());
        Assert.Equal(new string?[] { "  value  ", null }, sheet.ReadRows<StringValue>(2, 2).Select(p => p.Value).ToArray());
        Assert.Equal(new string?[] { "value" }, sheet.ReadRows<StringValue>(1, 1).Select(p => p.Value).ToArray());
    }

    [Fact]
    public void ReadRows_IndexCountAfterReadingAllRows_RewindsToStartIndex()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();

        Assert.Equal(4, sheet.ReadRows<StringValue>().Count());
        Assert.Equal(new string?[] { "value", "  value  " }, sheet.ReadRows<StringValue>(1, 2).Select(p => p.Value).ToArray());
    }

    [Fact]
    public void ReadRows_EnumeratedTwice_ReturnsSameRows()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();

        var rows = sheet.ReadRows<StringValue>();
        var expected = new string?[] { "value", "  value  ", null, "value" };
        Assert.Equal(expected, rows.Select(p => p.Value).ToArray());
        Assert.Equal(expected, rows.Select(p => p.Value).ToArray());
        Assert.Equal(4, sheet.CurrentRowIndex);
    }

    [Fact]
    public void ReadRows_AnyThenToList_ReturnsEveryRow()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();

        var rows = sheet.ReadRows<StringValue>();
        Assert.True(rows.Any());
        Assert.Equal(new string?[] { "value", "  value  ", null, "value" }, rows.Select(p => p.Value).ToList());
    }

    [Fact]
    public void ReadRows_AfterReadRow_EachEnumerationStartsAfterThatRow()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        Assert.Equal("value", sheet.ReadRow<StringValue>().Value);

        // The rows not read when ReadRows was called, however the sheet moves afterwards.
        var rows = sheet.ReadRows<StringValue>();
        Assert.Equal("  value  ", sheet.ReadRow<StringValue>().Value);
        Assert.Equal(new string?[] { "  value  ", null, "value" }, rows.Select(p => p.Value).ToArray());
        Assert.Equal(new string?[] { "  value  ", null, "value" }, rows.Select(p => p.Value).ToArray());
    }

    [Fact]
    public void ReadRows_HasHeadingFalseEnumeratedTwice_ReturnsSameRows()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        importer.Configuration.RegisterClassMap<StringValueClassMapColumnIndex>();
        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;

        var rows = sheet.ReadRows<StringValue>();
        var expected = new string?[] { "Value", "value", "  value  ", null, "value" };
        Assert.Equal(expected, rows.Select(p => p.Value).ToArray());
        Assert.Equal(expected, rows.Select(p => p.Value).ToArray());
    }

    [Fact]
    public void ReadRows_EnumeratedAfterReadingAnotherSheet_ReturnsSameRows()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();
        var rows = sheet.ReadRows<StringValue>();
        Assert.Equal(4, rows.Count());

        // The importer's reader is shared, so reading another sheet moves it.
        Assert.NotEmpty(importer.ReadSheet(1).ReadRows<Dictionary<string, string>>().ToArray());
        Assert.Equal(new string?[] { "value", "  value  ", null, "value" }, rows.Select(p => p.Value).ToArray());
    }

    [Fact]
    public void ReadRows_IndexCountEnumeratedTwice_ReturnsSameRows()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        var sheet = importer.ReadSheet();

        var rows = sheet.ReadRows<StringValue>(2, 2);
        Assert.Equal(new string?[] { "  value  ", null }, rows.Select(p => p.Value).ToArray());
        Assert.Equal(new string?[] { "  value  ", null }, rows.Select(p => p.Value).ToArray());
    }

    [Fact]
    public void ReadRowResults_EnumeratedTwice_ReturnsSameRows()
    {
        using var importer = Helpers.GetCsvImporter("Value\n1\nx\n3\n");
        var sheet = importer.ReadSheet();

        var results = sheet.ReadRowResults<IntValue>();
        Assert.True(results.Any());
        Assert.Equal(new[] { true, false, true }, results.Select(r => r.IsValid).ToArray());
        Assert.Equal(new[] { 1, 2, 3 }, results.Select(r => r.RowIndex).ToArray());
    }

    private class IntValue
    {
        public int Value { get; set; }
    }

    private sealed class CountingSink : IExcelErrorSink
    {
        public int Count { get; private set; }

        public List<int> Rows { get; } = [];

        public void Add(in ExcelMappingError error)
        {
            Count++;
            Rows.Add(error.RowIndex);
        }
    }

    public class SinkRow
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }

    [Fact]
    public void ReadRows_Sink_ReturnsValidRowsAndReportsOthers()
    {
        using var importer = Helpers.GetCsvImporter("Name,Age\nAlice,30\nBob,x\nCarol,\nDave,40\n");
        var sink = new CountingSink();

        var rows = importer.ReadSheet().ReadRows<SinkRow>(sink).ToArray();

        Assert.Equal(["Alice", "Dave"], rows.Select(r => r.Name));
        Assert.Equal([2, 3], sink.Rows);
    }

    [Fact]
    public void ReadRows_SinkRowWithSeveralErrors_ReportsEach()
    {
        using var importer = Helpers.GetCsvImporter("Name,Age,Other\nAlice,x,1\n");
        importer.Configuration.RegisterClassMap<SinkRow>(c =>
        {
            c.Map(p => p.Name).WithConverter(_ => throw new FormatException("bad name"));
            c.Map(p => p.Age);
        });
        var sink = new CountingSink();

        Assert.Empty(importer.ReadSheet().ReadRows<SinkRow>(sink));
        Assert.Equal(2, sink.Count);
    }

    [Fact]
    public void ReadRows_SinkEnumeratedTwice_ReportsAgain()
    {
        using var importer = Helpers.GetCsvImporter("Name,Age\nAlice,30\nBob,x\n");
        var sink = new CountingSink();

        var rows = importer.ReadSheet().ReadRows<SinkRow>(sink);
        Assert.Single(rows);
        Assert.Single(rows);
        Assert.Equal(2, sink.Count);
    }

    [Fact]
    public void ReadRows_SinkClassMap_UsesClassMap()
    {
        using var importer = Helpers.GetCsvImporter("Full Name,Age\nAlice,30\nBob,x\n");
        var map = new ExcelClassMap<SinkRow>();
        map.Map(p => p.Name).WithColumnName("Full Name");
        map.Map(p => p.Age);
        var sink = new CountingSink();

        Assert.Equal("Alice", Assert.Single(importer.ReadSheet().ReadRows(map, sink)).Name);
        Assert.Equal(1, sink.Count);
    }

    [Fact]
    public void ReadRows_SinkMissingColumn_ThrowsWhenCalled()
    {
        using var importer = Helpers.GetCsvImporter("Name\nAlice\n");
        var sheet = importer.ReadSheet();
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<SinkRow>(new CountingSink()));
    }

    [Fact]
    public void ReadRows_NullSink_ThrowsArgumentNullException()
    {
        using var importer = Helpers.GetCsvImporter("Name,Age\nAlice,30\n");
        var sheet = importer.ReadSheet();
        Assert.Throws<ArgumentNullException>("errors", () => sheet.ReadRows<SinkRow>((IExcelErrorSink)null!));
        Assert.Throws<ArgumentNullException>("errors", () => sheet.ReadRows(new ExcelClassMap<SinkRow>(), null!));
        Assert.Throws<ArgumentNullException>("classMap", () => sheet.ReadRows<SinkRow>(null!, new CountingSink()));
    }

    [Fact]
    public void ReadRows_SinkInvalidCells_DoesNotAllocatePerError()
    {
        // An invalid cell reported to a sink creates no exception: reading invalid cells
        // allocates about the same as reading valid ones.
        static long Allocated(string cell)
        {
            var csv = "Name,Age\n" + string.Concat(Enumerable.Range(0, 2000).Select(i => $"n,{cell}\n"));
            using var importer = Helpers.GetCsvImporter(csv);
            var sheet = importer.ReadSheet();
            var sink = new CountingSink();
            var rows = sheet.ReadRows<SinkRow>(sink);
            var before = GC.GetAllocatedBytesForCurrentThread();
            foreach (var row in rows)
            {
            }

            return GC.GetAllocatedBytesForCurrentThread() - before;
        }

        Allocated("1");
        Allocated("x");
        var valid = Allocated("1");
        var invalid = Allocated("x");

        // The valid rows allocate their objects; the invalid ones record their row numbers.
        Assert.True(invalid <= valid + (2000 * 16), $"Invalid cells allocated {invalid} bytes, valid cells {valid}.");
    }

    [Fact]
    public void Visibility_Csv_ReturnsVisible()
    {
        using var stream = Helpers.GetResource("Primitives.csv");
        using var importer = new ExcelImporter(stream, ExcelImporterFileType.Csv);
        var sheet = importer.ReadSheet();

        Assert.Equal(ExcelSheetVisibility.Visible, sheet.Visibility);
    }

    private class NameValueRow
    {
        public string? Name { get; set; }
        public string? Value { get; set; }
    }

    [Fact]
    public void ReadRows_CustomDataRangeColumnsStartAutoMapped_ReadsNamedColumns()
    {
        using var importer = Helpers.GetCsvImporter("Junk,Name,Value\nj1,n1,v1\nj2,n2,v2\n");
        var sheet = importer.ReadSheet();
        sheet.DataRange = new ExcelRange(Range.All, 1..);

        var rows = sheet.ReadRows<NameValueRow>().ToArray();
        Assert.Equal(["n1", "n2"], rows.Select(r => r.Name));
        Assert.Equal(["v1", "v2"], rows.Select(r => r.Value));
    }

    [Fact]
    public void ReadRows_CustomDataRangeColumnsStartInvalidValue_ReportsCorrectColumn()
    {
        using var importer = Helpers.GetCsvImporter("Junk,Value\n1,invalid\n");
        var sheet = importer.ReadSheet();
        sheet.DataRange = new ExcelRange(Range.All, 1..);

        var ex = Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<IntValueRow>().ToArray());
        Assert.Equal(1, ex.ColumnIndex);
        Assert.Contains("\"invalid\"", ex.Message);
        Assert.Contains("column \"Value\"", ex.Message);
    }

    private class IntValueRow
    {
        public int Value { get; set; }
    }

    [Fact]
    public void ReadRows_NoHeadingCustomDataRangeRowsStart_SkipsPrecedingRows()
    {
        using var importer = Helpers.GetCsvImporter("skip0\nskip1\nkeep1\nkeep2\n");
        importer.Configuration.RegisterClassMap<NameValueRow>(c => c.Map(p => p.Name).WithColumnIndex(0));
        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;
        sheet.DataRange = new ExcelRange(2.., Range.All);

        var rows = sheet.ReadRows<NameValueRow>().ToArray();
        Assert.Equal(["keep1", "keep2"], rows.Select(r => r.Name));
    }

    [Fact]
    public void ReadRows_NoHeadingCustomDataRangeRowsStartPastEnd_ReturnsEmpty()
    {
        using var importer = Helpers.GetCsvImporter("a\nb\n");
        importer.Configuration.RegisterClassMap<NameValueRow>(c => c.Map(p => p.Name).WithColumnIndex(0));
        var sheet = importer.ReadSheet();
        sheet.HasHeading = false;
        sheet.DataRange = new ExcelRange(5.., Range.All);

        Assert.Empty(sheet.ReadRows<NameValueRow>());
    }

    [Fact]
    public void ReadRows_SkipBlankLinesCustomDataRangeColumns_IgnoresColumnsOutsideRange()
    {
        using var importer = Helpers.GetCsvImporter("Name,Value,Note\nn1,v1,\n,,note\nn2,v2,\n");
        importer.Configuration.SkipBlankLines = true;
        var sheet = importer.ReadSheet();
        sheet.DataRange = new ExcelRange(Range.All, 0..2);

        var rows = sheet.ReadRows<NameValueRow>().ToArray();
        Assert.Equal(["n1", "n2"], rows.Select(r => r.Name));
    }

    private class MatchingRow
    {
        public string[] Values { get; set; } = default!;
    }

    [Fact]
    public void ReadRows_ColumnsMatchingCustomDataRangeColumns_MatchesColumnsInRange()
    {
        using var importer = Helpers.GetCsvImporter("A1,Skip,A2,A3\n1,x,2,3\n");
        importer.Configuration.RegisterClassMap<MatchingRow>(c => c
            .Map(p => p.Values)
            .WithColumnsMatching(n => n.StartsWith("A")));
        var sheet = importer.ReadSheet();
        sheet.DataRange = new ExcelRange(Range.All, 1..3);

        var row = sheet.ReadRows<MatchingRow>().Single();
        Assert.Equal(["2"], row.Values);
    }

    // A class map given to a reading method, rather than one registered or auto-mapped: it reads
    // the "Other" column, which auto-mapping MapRow wouldn't.
    private static ExcelClassMap<MapRow> CreateGivenMap()
    {
        var classMap = new ExcelClassMap<MapRow>();
        classMap.Map(r => r.Name).WithColumnName("Other");
        return classMap;
    }

    public class MapRow
    {
        public string? Name { get; set; }
    }

    private const string MapCsv = "Name,Other\na,x\nb,y\n";

    [Fact]
    public void ReadRows_ClassMap_UsesClassMap()
    {
        using var importer = Helpers.GetCsvImporter(MapCsv);
        importer.Configuration.RegisterClassMap<MapRow>(c => c.Map(r => r.Name));
        var sheet = importer.ReadSheet();
        Assert.Equal(["x", "y"], sheet.ReadRows(CreateGivenMap()).Select(r => r.Name));
    }

    [Fact]
    public void ReadRows_ClassMapRange_UsesClassMap()
    {
        using var importer = Helpers.GetCsvImporter(MapCsv);
        var sheet = importer.ReadSheet();
        Assert.Equal(["y"], sheet.ReadRows(CreateGivenMap(), 2, 1).Select(r => r.Name));
    }

    [Fact]
    public void ReadRow_ClassMap_UsesClassMap()
    {
        using var importer = Helpers.GetCsvImporter(MapCsv);
        var sheet = importer.ReadSheet();
        var classMap = CreateGivenMap();
        Assert.Equal("x", sheet.ReadRow(classMap).Name);
        Assert.True(sheet.TryReadRow(classMap, out var row));
        Assert.Equal("y", row.Name);
        Assert.False(sheet.TryReadRow(classMap, out _));
        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow(classMap));
    }

    [Fact]
    public void ReadRowResults_ClassMap_UsesClassMap()
    {
        using var importer = Helpers.GetCsvImporter(MapCsv);
        var results = importer.ReadSheet().ReadRowResults(CreateGivenMap()).ToArray();
        Assert.Equal(["x", "y"], results.Select(r => r.Value!.Name));
    }

    [Fact]
    public void ValidateHeading_ClassMap_ChecksClassMapsColumns()
    {
        using var importer = Helpers.GetCsvImporter("Name\na\n");
        var sheet = importer.ReadSheet();
        Assert.False(sheet.TryValidateHeading(CreateGivenMap(), out var problems));
        Assert.Single(problems);
        Assert.Throws<ExcelMappingException>(() => sheet.ValidateHeading(CreateGivenMap()));

        using var valid = Helpers.GetCsvImporter(MapCsv);
        var validSheet = valid.ReadSheet();
        validSheet.ValidateHeading(CreateGivenMap());
        Assert.True(validSheet.TryValidateHeading(CreateGivenMap(), out problems));
        Assert.Empty(problems);
    }

    [Fact]
    public void FindHeading_ClassMap_FindsClassMapsColumns()
    {
        const string Csv = "Title,\nName,Other\na,x\n";
        using (var importer = Helpers.GetCsvImporter(Csv))
        {
            var sheet = importer.ReadSheet();
            Assert.Equal(["Name", "Other"], sheet.FindHeading(CreateGivenMap()).ColumnNames);
            Assert.Equal("x", sheet.ReadRow(CreateGivenMap()).Name);
        }

        using (var importer = Helpers.GetCsvImporter(Csv))
        {
            Assert.Equal(1, importer.ReadSheet().FindHeading(CreateGivenMap(), 5).ColumnNames.Count - 1);
        }

        using (var importer = Helpers.GetCsvImporter(Csv))
        {
            Assert.True(importer.ReadSheet().TryFindHeading(CreateGivenMap(), out var heading));
            Assert.Equal(2, heading.ColumnNames.Count);
        }

        using (var importer = Helpers.GetCsvImporter(Csv))
        {
            Assert.False(importer.ReadSheet().TryFindHeading(CreateGivenMap(), 1, out _));
        }
    }

    [Fact]
    public void ReadRows_NullClassMap_ThrowsArgumentNullException()
    {
        using var importer = Helpers.GetCsvImporter(MapCsv);
        var sheet = importer.ReadSheet();
        Assert.Throws<ArgumentNullException>("classMap", () => sheet.ReadRows<MapRow>((ExcelClassMap<MapRow>)null!));
        Assert.Throws<ArgumentNullException>("classMap", () => sheet.ReadRows<MapRow>(null!, 1, 1));
        Assert.Throws<ArgumentNullException>("classMap", () => sheet.ReadRow<MapRow>(null!));
        Assert.Throws<ArgumentNullException>("classMap", () => sheet.TryReadRow<MapRow>(null!, out _));
        Assert.Throws<ArgumentNullException>("classMap", () => sheet.ReadRowResults<MapRow>(null!));
        Assert.Throws<ArgumentNullException>("classMap", () => sheet.ValidateHeading<MapRow>(null!));
        Assert.Throws<ArgumentNullException>("classMap", () => sheet.TryValidateHeading<MapRow>(null!, out _));
        Assert.Throws<ArgumentNullException>("classMap", () => sheet.FindHeading<MapRow>(null!));
        Assert.Throws<ArgumentNullException>("classMap", () => sheet.FindHeading<MapRow>(null!, 1));
        Assert.Throws<ArgumentNullException>("classMap", () => sheet.TryFindHeading<MapRow>(null!, out _));
        Assert.Throws<ArgumentNullException>("classMap", () => sheet.TryFindHeading<MapRow>(null!, 1, out _));
    }
}

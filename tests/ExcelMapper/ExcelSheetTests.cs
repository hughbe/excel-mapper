using System.Linq;

namespace ExcelMapper.Tests;

public class ExcelSheetTests
{
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

        StringValue[] rows2 = [.. sheet.ReadRows<StringValue>()];
        Assert.Empty(rows2.Select(p => p.Value).ToArray());
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

        StringValue[] rows2 = [.. sheet.ReadRows<StringValue>()];
        Assert.Empty(rows2.Select(p => p.Value).ToArray());
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
        Assert.Equal(999, sheet.CurrentRowIndex);
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

        Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<BlankLinesClass>());
        Assert.Equal(10, sheet.CurrentRowIndex);
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
        Assert.Equal(999, sheet.CurrentRowIndex);
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
        Assert.Equal(new DateTime(2025, 10, 25, 0, 0, 0).ToString(), row5.StringValue);
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
        Assert.Equal("No more rows in sheet \"Primitives\".", ex.Message);
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
}

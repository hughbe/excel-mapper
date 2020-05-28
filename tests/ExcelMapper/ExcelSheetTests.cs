using System;
using System.Collections.Generic;
using System.Linq;
using Xunit;

namespace ExcelMapper.Tests
{
    public class ExcelSheetTests
    {
        [Fact]
        public void Visibility_Get_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("HiddenSheets.xlsx");
            ExcelSheet sheet = importer.ReadSheet();
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
            ExcelSheet sheet = importer.ReadSheet();
            ExcelHeading heading = sheet.ReadHeading();
            Assert.Same(heading, sheet.Heading);
            Assert.Equal(new string[] { "Int Value", "StringValue", "Bool Value", "Enum Value", "DateValue", "ArrayValue", "MappedValue", "TrimmedValue" }, heading.ColumnNames);
            Assert.Equal(-1, sheet.CurrentRowIndex);
        }

        [Fact]
        public void ReadHeading_EmptyColumnName_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("EmptyColumns.xlsx");
            ExcelSheet sheet = importer.ReadSheet();
            ExcelHeading heading = sheet.ReadHeading();
            Assert.Same(heading, sheet.Heading);
            Assert.Equal(new string[] { "", "Column2", "", " Column4 " }, heading.ColumnNames);
            Assert.Equal(-1, sheet.CurrentRowIndex);
        }

        [Fact]
        public void ReadHeading_NonZeroHeadingIndex_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Non Zero Header Index.xlsx");
            ExcelSheet sheet = importer.ReadSheet();
            sheet.HeadingIndex = 3;

            ExcelHeading heading = sheet.ReadHeading();
            Assert.Same(heading, sheet.Heading);
            Assert.Equal(new string[] { "Value" }, heading.ColumnNames);
            Assert.Equal(-1, sheet.CurrentRowIndex);
        }

        [Fact]
        public void ReadHeading_AlreadyReadHeading_ThrowsExcelMappingException()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            Assert.Throws<ExcelMappingException>(() => sheet.ReadHeading());
        }

        [Fact]
        public void ReadHeading_DoesNotHaveHasHeading_ThrowsExcelMappingException()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            ExcelSheet sheet = importer.ReadSheet();
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
        public void ReadRows_NotReadHeading_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Strings.xlsx");
            ExcelSheet sheet = importer.ReadSheet();

            IEnumerable<StringValue> rows = sheet.ReadRows<StringValue>();
            Assert.Equal(new string[] { "value", "  value  ", null, "value" }, rows.Select(p => p.Value).ToArray());
            Assert.Equal(3, sheet.CurrentRowIndex);

            Assert.NotNull(sheet.Heading);
            Assert.True(sheet.HasHeading);
        }

        [Fact]
        public void ReadRows_HasHeadingFalse_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Strings.xlsx");
            importer.Configuration.RegisterClassMap<StringValueClassMapColumnIndex>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.HasHeading = false;

            IEnumerable<StringValue> rows = sheet.ReadRows<StringValue>();
            Assert.Equal(new string[] { "Value", "value", "  value  ", null, "value" }, rows.Select(p => p.Value).ToArray());
            Assert.Equal(4, sheet.CurrentRowIndex);

            Assert.Null(sheet.Heading);
            Assert.False(sheet.HasHeading);
        }

        [Fact]
        public void ReadRows_ReadHeading_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Strings.xlsx");
            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            IEnumerable<StringValue> rows = sheet.ReadRows<StringValue>();
            Assert.Equal(new string[] { "value", "  value  ", null, "value"  }, rows.Select(p => p.Value).ToArray());
            Assert.Equal(3, sheet.CurrentRowIndex);

            Assert.NotNull(sheet.Heading);
            Assert.True(sheet.HasHeading);
        }

        [Fact]
        public void ReadRows_ReadHeadingNonZeroHeadingIndex_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Non Zero Header Index.xlsx");
            ExcelSheet sheet = importer.ReadSheet();
            sheet.HeadingIndex = 3;
            sheet.ReadHeading();

            IEnumerable<StringValue> rows = sheet.ReadRows<StringValue>();
            Assert.Equal(new string[] { "value", "  value  ", null, "value" }, rows.Select(p => p.Value).ToArray());
            Assert.Equal(3, sheet.CurrentRowIndex);

            Assert.NotNull(sheet.Heading);
            Assert.True(sheet.HasHeading);
        }

        [Fact]
        public void ReadRows_AllReadHasHeadingTrue_ReturnsEmpty()
        {
            using var importer = Helpers.GetImporter("Strings.xlsx");
            ExcelSheet sheet = importer.ReadSheet();

            IEnumerable<StringValue> rows1 = sheet.ReadRows<StringValue>();
            Assert.Equal(new string[] { "value", "  value  ", null, "value" }, rows1.Select(p => p.Value).ToArray());
            Assert.Equal(3, sheet.CurrentRowIndex);

            Assert.NotNull(sheet.Heading);
            Assert.True(sheet.HasHeading);

            StringValue[] rows2 = sheet.ReadRows<StringValue>().ToArray();
            Assert.Empty(rows2.Select(p => p.Value).ToArray());
        }

        [Fact]
        public void ReadRows_AllReadingHasHeadingFalse_ReturnsEmpty()
        {
            using var importer = Helpers.GetImporter("Strings.xlsx");
            importer.Configuration.RegisterClassMap<StringValueClassMapColumnIndex>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.HasHeading = false;

            IEnumerable<StringValue> rows1 = sheet.ReadRows<StringValue>();
            Assert.Equal(new string[] { "Value", "value", "  value  ", null, "value" }, rows1.Select(p => p.Value).ToArray());
            Assert.Equal(4, sheet.CurrentRowIndex);

            Assert.Null(sheet.Heading);
            Assert.False(sheet.HasHeading);

            StringValue[] rows2 = sheet.ReadRows<StringValue>().ToArray();
            Assert.Empty(rows2.Select(p => p.Value).ToArray());
        }

        [Fact]
        public void ReadRows_EmptySheetNoHeading_ReturnsEmpty()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            importer.Configuration.RegisterClassMap<StringValueClassMapColumnIndex>();

            importer.ReadSheet();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.HasHeading = false;

            Assert.Empty(sheet.ReadRows<StringValue>());
        }

        public static IEnumerable<object[]> ReadRows_Area_TestData()
        {
            yield return new object[] { 1, 2, new string[] { "value", "  value  " } };
            yield return new object[] { 0, 4, new string[] { "value", "  value  ", null, "value" } };
            yield return new object[] { 1, 0, new string[0] };
        }

        [Theory]
        [MemberData(nameof(ReadRows_Area_TestData))]
        public void ReadRows_IndexCount_ReturnsExpected(int startIndex, int count, string[] expectedValues)
        {
            using var importer = Helpers.GetImporter("Strings.xlsx");
            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            IEnumerable<StringValue> rows = sheet.ReadRows<StringValue>(startIndex, count);
            Assert.Equal(expectedValues, rows.Select(p => p.Value).ToArray());
            Assert.Equal(startIndex + count, sheet.CurrentRowIndex);

            Assert.NotNull(sheet.Heading);
            Assert.True(sheet.HasHeading);
        }

        [Fact]
        public void ReadRows_BlankLinesNotSkipped_ThrowsExcelMappingException()
        {
            using var importer = Helpers.GetImporter("BlankLines.xlsx");
            ExcelSheet sheet = importer.ReadSheet();

            Assert.Throws<ExcelMappingException>(() => sheet.ReadRows<BlankLinesClass>().ToArray());
            Assert.Equal(0, sheet.CurrentRowIndex);
        }

        [Fact]
        public void ReadRows_BlankLinesSkipped_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("BlankLines.xlsx");
            importer.Configuration.SkipBlankLines = true;
            ExcelSheet sheet = importer.ReadSheet();

            BlankLinesClass[] rows = sheet.ReadRows<BlankLinesClass>().ToArray();
            Assert.Equal(4, rows.Length);
            Assert.Equal("A", rows[0].StringValue);
            Assert.Equal(1, rows[0].IntValue);
            Assert.Equal("B", rows[1].StringValue);
            Assert.Equal(2, rows[1].IntValue);
            Assert.Null(rows[2].StringValue);
            Assert.Equal(3, rows[2].IntValue);
            Assert.Equal("C", rows[3].StringValue);
            Assert.Equal(0, rows[3].IntValue);
            Assert.Equal(998, sheet.CurrentRowIndex);
        }

        [Fact]
        public void ReadRow_BlankLinesNotSkipped_ThrowsExcelMappingException()
        {
            using var importer = Helpers.GetImporter("BlankLines.xlsx");
            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<BlankLinesClass>());
            Assert.Equal(0, sheet.CurrentRowIndex);

            BlankLinesClass row1 = sheet.ReadRow<BlankLinesClass>();
            Assert.Equal("A", row1.StringValue);
            Assert.Equal(1, row1.IntValue);
            Assert.Equal(1, sheet.CurrentRowIndex);

            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<BlankLinesClass>());
            Assert.Equal(2, sheet.CurrentRowIndex);

            BlankLinesClass row2 = sheet.ReadRow<BlankLinesClass>();
            Assert.Equal("B", row2.StringValue);
            Assert.Equal(2, row2.IntValue);
            Assert.Equal(3, sheet.CurrentRowIndex);

            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<BlankLinesClass>());
            Assert.Equal(4, sheet.CurrentRowIndex);

            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<BlankLinesClass>());
            Assert.Equal(5, sheet.CurrentRowIndex);

            BlankLinesClass row3 = sheet.ReadRow<BlankLinesClass>();
            Assert.Null(row3.StringValue);
            Assert.Equal(3, row3.IntValue);
            Assert.Equal(6, sheet.CurrentRowIndex);

            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<BlankLinesClass>());
            Assert.Equal(7, sheet.CurrentRowIndex);

            BlankLinesClass row4 = sheet.ReadRow<BlankLinesClass>();
            Assert.Equal("C", row4.StringValue);
            Assert.Equal(0, row4.IntValue);
            Assert.Equal(8, sheet.CurrentRowIndex);

            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<BlankLinesClass>());
            Assert.Equal(9, sheet.CurrentRowIndex);
        }

        [Fact]
        public void ReadRow_BlankLinesSkipped_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("BlankLines.xlsx");
            importer.Configuration.SkipBlankLines = true;
            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            BlankLinesClass row1 = sheet.ReadRow<BlankLinesClass>();
            Assert.Equal("A", row1.StringValue);
            Assert.Equal(1, row1.IntValue);
            Assert.Equal(1, sheet.CurrentRowIndex);

            BlankLinesClass row2 = sheet.ReadRow<BlankLinesClass>();
            Assert.Equal("B", row2.StringValue);
            Assert.Equal(2, row2.IntValue);
            Assert.Equal(3, sheet.CurrentRowIndex);

            BlankLinesClass row3 = sheet.ReadRow<BlankLinesClass>();
            Assert.Null(row3.StringValue);
            Assert.Equal(3, row3.IntValue);
            Assert.Equal(6, sheet.CurrentRowIndex);

            BlankLinesClass row4 = sheet.ReadRow<BlankLinesClass>();
            Assert.Equal("C", row4.StringValue);
            Assert.Equal(0, row4.IntValue);
            Assert.Equal(8, sheet.CurrentRowIndex);

            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<BlankLinesClass>());
            Assert.Equal(998, sheet.CurrentRowIndex);
        }

        [Fact]
        public void ReadRows_NegativeStartIndex_ThrowsArgumentOutOfRangeException()
        {
            using var importer = Helpers.GetImporter("Strings.xlsx");
            ExcelSheet sheet = importer.ReadSheet();
            Assert.Throws<ArgumentOutOfRangeException>("startIndex", () => sheet.ReadRows<StringValue>(-1, 0).ToArray());
        }

        [Fact]
        public void ReadRows_NegativeCount_ThrowsArgumentOutOfRangeException()
        {
            using var importer = Helpers.GetImporter("Strings.xlsx");
            ExcelSheet sheet = importer.ReadSheet();
            Assert.Throws<ArgumentOutOfRangeException>("count", () => sheet.ReadRows<StringValue>(0, -1).ToArray());
        }

        [Fact]
        public void ReadRow_HasHeadingFalse_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Strings.xlsx");
            importer.Configuration.RegisterClassMap<StringValueClassMapColumnIndex>();

            ExcelSheet sheet = importer.ReadSheet();
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
            ExcelSheet sheet = importer.ReadSheet();
            sheet.HasHeading = false;

            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<StringValue>());
        }

        [Fact]
        public void ReadRow_HasHeadingFalseColumnNameMapping_ThrowsExcelMappingException()
        {
            using var importer = Helpers.GetImporter("Strings.xlsx");
            importer.Configuration.RegisterClassMap<StringValueClassMapColumnName>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.HasHeading = false;

            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<StringValue>());
        }

        [Fact]
        public void ReadRow_HasHeadingFalseColumnNamesMapping_ThrowsExcelMappingException()
        {
            using var importer = Helpers.GetImporter("Strings.xlsx");
            importer.Configuration.RegisterClassMap<StringValuesClassMapColumnNames>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.HasHeading = false;

            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<StringValues>());
        }

        [Fact]
        public void HasHeading_SetWhenAlreadyRead_InvalidOperationException()
        {
            using var importer = Helpers.GetImporter("Strings.xlsx");
            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            Assert.Throws<InvalidOperationException>(() => sheet.HasHeading = false);
            Assert.Throws<InvalidOperationException>(() => sheet.HasHeading = true);
        }

        [Fact]
        public void HeadingIndex_SetNegative_ThrowsArgumentOutOfRangeException()
        {
            using var importer = Helpers.GetImporter("Strings.xlsx");
            ExcelSheet sheet = importer.ReadSheet();
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
            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            Assert.Throws<InvalidOperationException>(() => sheet.HeadingIndex = 0);
        }

        [Fact]
        public void HeadingIndex_SetWhenHasHeadingFalse_ThrowsInvalidOperationException()
        {
            using var importer = Helpers.GetImporter("Strings.xlsx");
            ExcelSheet sheet = importer.ReadSheet();
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
            public string Value { get; set; }
        }

        private class StringValues
        {
            public string[] Value { get; set; }
        }

        [Fact]
        public void ReadRow_CantMapType_ThrowsExcelMappingException()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            ExcelSheet sheet = importer.ReadSheet();
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Helpers.IListInterface>());
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IConvertible>());
        }

        [Fact]
        public void ReadHeading_TooLargeHeadingIndex_ThrowsExcelMappingException()
        {
            using var importer = Helpers.GetImporter("Non Zero Header Index.xlsx");
            ExcelSheet sheet = importer.ReadSheet();
            sheet.HeadingIndex = 8;

            Assert.Throws<ExcelMappingException>(() => sheet.ReadHeading());
        }

        [Fact]
        public void ReadRow_NoMoreRows_ThrowsExcelMappingException()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            ExcelSheet sheet = importer.ReadSheet();
            Assert.NotEmpty(sheet.ReadRows<object>().ToArray());

            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<object>());
            Assert.False(sheet.TryReadRow(out object row));
            Assert.Null(row);
        }

        private class BlankLinesClass
        {
            public string StringValue { get; set; }
            public int IntValue { get; set; }
        }
    }
}

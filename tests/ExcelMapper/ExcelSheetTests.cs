using System;
using System.Collections.Generic;
using System.Linq;
using Xunit;

namespace ExcelMapper.Tests
{
    public class ExcelSheetTests
    {
        [Fact]
        public void ReadHeading_HasHeading_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                ExcelSheet sheet = importer.ReadSheet();
                ExcelHeading heading = sheet.ReadHeading();
                Assert.Same(heading, sheet.Heading);

                Assert.Equal(new string[] { "Int Value", "StringValue", "Bool Value", "Enum Value", "DateValue", "ArrayValue", "MappedValue", "TrimmedValue" }, heading.ColumnNames);
            }
        }

        [Fact]
        public void ReadHeading_EmptyColumnName_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("EmptyColumns.xlsx"))
            {
                ExcelSheet sheet = importer.ReadSheet();
                ExcelHeading heading = sheet.ReadHeading();
                Assert.Same(heading, sheet.Heading);

                Assert.Equal(new string[] { "", "Column2", "", " Column4 " }, heading.ColumnNames);
            }
        }

        [Fact]
        public void ReadHeading_AlreadyReadHeading_ThrowsExcelMappingException()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                Assert.Throws<ExcelMappingException>(() => sheet.ReadHeading());
            }
        }

        [Fact]
        public void ReadHeading_DoesNotHaveHasHeading_ThrowsExcelMappingException()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                ExcelSheet sheet = importer.ReadSheet();
                sheet.HasHeading = false;

                Assert.Throws<ExcelMappingException>(() => sheet.ReadHeading());
            }
        }

        [Fact]
        public void ReadHeading_NoRows_ThrowsExcelMappingException()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                importer.ReadSheet();

                ExcelSheet emptySheet = importer.ReadSheet();
                Assert.Throws<ExcelMappingException>(() => emptySheet.ReadHeading());
            }
        }

        [Fact]
        public void ReadRows_NotReadHeading_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("Strings.xlsx"))
            {
                ExcelSheet sheet = importer.ReadSheet();

                IEnumerable<StringValue> rows = sheet.ReadRows<StringValue>();
                Assert.Equal(new string[] { "value", "  value  ", null, "value" }, rows.Select(p => p.Value).ToArray());

                Assert.NotNull(sheet.Heading);
                Assert.True(sheet.HasHeading);
            }
        }

        [Fact]
        public void ReadRows_HasHeadingFalse_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("Strings.xlsx"))
            {
                importer.Configuration.RegisterClassMap<StringValueClassMap>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.HasHeading = false;

                IEnumerable<StringValue> rows = sheet.ReadRows<StringValue>();
                Assert.Equal(new string[] { "Value", "value", "  value  ", null, "value" }, rows.Select(p => p.Value).ToArray());

                Assert.Null(sheet.Heading);
                Assert.False(sheet.HasHeading);
            }
        }

        [Fact]
        public void ReadRows_ReadHeading_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("Strings.xlsx"))
            {
                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                IEnumerable<StringValue> rows = sheet.ReadRows<StringValue>();
                Assert.Equal(new string[] { "value", "  value  ", null, "value"  }, rows.Select(p => p.Value).ToArray());

                Assert.NotNull(sheet.Heading);
                Assert.True(sheet.HasHeading);
            }
        }

        [Fact]
        public void ReadRow_ReadHasHeadingFalse_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("Strings.xlsx"))
            {
                ExcelSheet sheet = importer.ReadSheet();
                sheet.HasHeading = false;

                StringValue value = sheet.ReadRow<StringValue>();
                Assert.Equal("Value", value.Value);

                Assert.Null(sheet.Heading);
                Assert.False(sheet.HasHeading);
            }
        }

        private class StringValueClassMap : ExcelClassMap<StringValue>
        {
            public StringValueClassMap()
            {
                Map(value => value.Value)
                    .WithColumnIndex(0);
            }
        }

        private class StringValue
        {
            public string Value { get; set; }
        }

        [Fact]
        public void ReadRow_CantMapType_ThrowsExcelMappingException()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                ExcelSheet sheet = importer.ReadSheet();
                Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Helpers.IListInterface>());
                Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IConvertible>());
            }
        }

        [Fact]
        public void ReadRow_NoMoreRows_ThrowsExcelMappingException()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                ExcelSheet sheet = importer.ReadSheet();
                Assert.NotEmpty(sheet.ReadRows<object>().ToArray());

                Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<object>());
                Assert.False(sheet.TryReadRow(out object row));
                Assert.Null(row);
            }
        }
    }
}

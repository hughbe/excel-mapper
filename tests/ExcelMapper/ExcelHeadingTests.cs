using System;
using System.Linq;
using Xunit;

namespace ExcelMapper.Tests
{
    public class ExcelHeadingTests
    {
        [Fact]
        public void GetColumnName_GetColumnIndex_Roundtrips()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            ExcelSheet sheet = importer.ReadSheet();
            ExcelHeading heading = sheet.ReadHeading();

            string[] columnNames = heading.ColumnNames.ToArray();
            for (int i = 0; i < columnNames.Length; i++)
            {
                Assert.Equal(i, heading.GetColumnIndex(columnNames[i]));
                Assert.Equal(columnNames[i], heading.GetColumnName(i));
            }
        }

        [Fact]
        public void GetColumnIndex_InvokeValidColumnName_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            ExcelSheet sheet = importer.ReadSheet();
            ExcelHeading heading = sheet.ReadHeading();

            Assert.Equal(1, heading.GetColumnIndex("StringValue"));
            Assert.Equal(1, heading.GetColumnIndex("stringvalue"));
        }

        [Fact]
        public void GetColumnIndex_NullColumnName_ThrowsArgumentNullException()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            ExcelSheet sheet = importer.ReadSheet();
            ExcelHeading heading = sheet.ReadHeading();

            Assert.Throws<ArgumentNullException>("key", () => heading.GetColumnIndex(null));
        }

        [Theory]
        [InlineData("")]
        [InlineData("NoSuchColumn")]
        public void GetColumnIndex_NoSuchColumnName_ThrowsExcelMappingException(string columnName)
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            ExcelSheet sheet = importer.ReadSheet();
            ExcelHeading heading = sheet.ReadHeading();

            Assert.Throws<ExcelMappingException>(() => heading.GetColumnIndex(columnName));
        }

        [Fact]
        public void GetColumnName_GetFirstColumnMatchingIndex_Roundtrips()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            ExcelSheet sheet = importer.ReadSheet();
            ExcelHeading heading = sheet.ReadHeading();

            string[] columnNames = heading.ColumnNames.ToArray();
            for (int i = 0; i < columnNames.Length; i++)
            {
                Assert.Equal(i, heading.GetFirstColumnMatchingIndex(e => e == columnNames[i]));
                Assert.Equal(columnNames[i], heading.GetColumnName(i));
            }
        }

        [Theory]
        [InlineData("")]
        [InlineData("NoSuchColumn")]
        public void GetFirstColumnMatchingIndex_NoSuchColumnName_ThrowsExcelMappingException(string columnName)
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            ExcelSheet sheet = importer.ReadSheet();
            ExcelHeading heading = sheet.ReadHeading();

            Assert.Throws<ExcelMappingException>(() => heading.GetFirstColumnMatchingIndex(e => e == columnName));
        }

        [Fact]
        public void GetColumnName_TryGetColumnIndex_Roundtrips()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            ExcelSheet sheet = importer.ReadSheet();
            ExcelHeading heading = sheet.ReadHeading();

            string[] columnNames = heading.ColumnNames.ToArray();
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
            ExcelSheet sheet = importer.ReadSheet();
            ExcelHeading heading = sheet.ReadHeading();

            Assert.True(heading.TryGetColumnIndex("StringValue", out int index));
            Assert.Equal(1, index);
            Assert.True(heading.TryGetColumnIndex("stringvalue", out index));
            Assert.Equal(1, index);
            Assert.Equal(1, index);
        }

        [Fact]
        public void TryGetColumnIndex_NullColumnName_ThrowsArgumentNullException()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            ExcelSheet sheet = importer.ReadSheet();
            ExcelHeading heading = sheet.ReadHeading();

            Assert.Throws<ArgumentNullException>("key", () => heading.TryGetColumnIndex(null, out int index));
        }

        [Theory]
        [InlineData("")]
        [InlineData("NoSuchColumn")]
        public void TryGetColumnIndex_NoSuchColumnName_ReturnsFalse(string columnName)
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            ExcelSheet sheet = importer.ReadSheet();
            ExcelHeading heading = sheet.ReadHeading();

            Assert.False(heading.TryGetColumnIndex(columnName, out int index));
            Assert.Equal(0, index);
        }

        [Fact]
        public void GetColumnName_TryGetFirstColumnMatchingIndex_Roundtrips()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            ExcelSheet sheet = importer.ReadSheet();
            ExcelHeading heading = sheet.ReadHeading();

            string[] columnNames = heading.ColumnNames.ToArray();
            for (int i = 0; i < columnNames.Length; i++)
            {
                Assert.True(heading.TryGetFirstColumnMatchingIndex(e => e == columnNames[i], out int index));
                Assert.Equal(i, index);
                Assert.Equal(columnNames[i], heading.GetColumnName(i));
            }
        }

        [Theory]
        [InlineData("")]
        [InlineData("NoSuchColumn")]
        public void TryGetFirstColumnMatchingIndex_NoSuchColumnName_ReturnsFalse(string columnName)
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            ExcelSheet sheet = importer.ReadSheet();
            ExcelHeading heading = sheet.ReadHeading();

            Assert.False(heading.TryGetFirstColumnMatchingIndex(e => e == columnName, out int index));
            Assert.Equal(-1, index);
        }
    }
}

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
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                ExcelSheet sheet = importer.ReadSheet();
                ExcelHeading heading = sheet.ReadHeading();

                string[] columnNames = heading.ColumnNames.ToArray();
                for (int i = 0; i < columnNames.Length; i++)
                {
                    Assert.Equal(i, heading.GetColumnIndex(columnNames[i]));
                    Assert.Equal(columnNames[i], heading.GetColumnName(i));
                }
            }
        }
        
        [Fact]
        public void GetColumnIndex_NullColumnName_ThrowsArgumentNullException()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                ExcelSheet sheet = importer.ReadSheet();
                ExcelHeading heading = sheet.ReadHeading();

                Assert.Throws<ArgumentNullException>("key", () => heading.GetColumnIndex(null));
            }
        }

        [Theory]
        [InlineData("")]
        [InlineData("NoSuchColumn")]
        public void GetColumnIndex_NoSuchColumnName_ThrowsExcelMappingException(string columnName)
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                ExcelSheet sheet = importer.ReadSheet();
                ExcelHeading heading = sheet.ReadHeading();

                Assert.Throws<ExcelMappingException>(() => heading.GetColumnIndex(columnName));
            }
        }
    }
}

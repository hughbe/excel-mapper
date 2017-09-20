using System.Linq;
using Xunit;

namespace ExcelMapper.Tests
{
    public class MapDuplicatedColumnsTests
    {
        [Fact]
        public void ReadHeading_WithDuplicatedColumns_AssignsRandomNameToTheSecondOne()
        {
            using (var importer = Helpers.GetImporter("DuplicatedColumns.xlsx"))
            {
                var sheet = importer.ReadSheet();
                var heading = sheet.ReadHeading();
                var columnNames = heading.ColumnNames.ToArray();

                Assert.Equal("MyColumn", columnNames[0]);
                Assert.StartsWith("MyColumn", columnNames[0]);
            }
        }

        [Fact]
        public void ReadRow_CustomMapped_CorrectlyMapByIndex()
        {
            using (var importer = Helpers.GetImporter("DuplicatedColumns.xlsx"))
            {
                importer.Configuration.RegisterClassMap<MyValuesMap>();

                var sheet = importer.ReadSheet();
                sheet.ReadHeading();

                var row1 = sheet.ReadRow<MyValues>();
                Assert.Equal("value1", row1.MyColumn);
                Assert.Equal("value2", row1.AnotherColumn);
            }
        }

        [Fact]
        public void ReadRow_AutoMapped_Success()
        {
            using (var importer = Helpers.GetImporter("DuplicatedColumns.xlsx"))
            {
                var sheet = importer.ReadSheet();
                sheet.ReadHeading();

                var row1 = sheet.ReadRow<MyValuesWithOneColumn>();
                Assert.Equal("value1", row1.MyColumn);
            }
        }

        private class MyValuesWithOneColumn
        {
            public string MyColumn { get; set; }
        }

        private class MyValues
        {
            public string MyColumn { get; set; }
            public string AnotherColumn { get; set; }
        }

        private class MyValuesMap : ExcelClassMap<MyValues>
        {
            public MyValuesMap()
            {
                Map(o => o.MyColumn).WithColumnName("MyColumn");
                Map(o => o.AnotherColumn).WithColumnIndex(1);
            }
        }
    }
}

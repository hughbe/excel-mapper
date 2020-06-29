using System;
using System.Collections.Generic;
using System.Dynamic;
using Xunit;

namespace ExcelMapper.Tests
{
    public class MapExpandoObjectTests
    {
        [Fact]
        public void ReadRow_AutoMappedExpandoObject_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            dynamic row1 = sheet.ReadRow<ExpandoObjectClass>().Value;
            Assert.Equal("a", row1.Column1);
            Assert.Equal("1", row1.Column2);
            Assert.Equal("2", row1.Column3);
            Assert.Null(row1.Column4);

            dynamic row2 = sheet.ReadRow<ExpandoObjectClass>().Value;
            Assert.Equal("b", row2.Column1);
            Assert.Equal("0", row2.Column2);
            Assert.Equal("0", row2.Column3);
            Assert.Null(row2.Column4);

            dynamic row3 = sheet.ReadRow<ExpandoObjectClass>().Value;
            Assert.Equal("c", row3.Column1);
            Assert.Equal("-2", row3.Column2);
            Assert.Equal("-1", row3.Column3);
            Assert.Null(row3.Column4);
        }

        [Fact]
        public void ReadRow_DefaultMappedExpandoObject_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
            importer.Configuration.RegisterClassMap<DefaultExpandoObjectClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            dynamic row1 = sheet.ReadRow<ExpandoObjectClass>().Value;
            Assert.Equal("a", row1.Column1);
            Assert.Equal("1", row1.Column2);
            Assert.Equal("2", row1.Column3);
            Assert.Null(row1.Column4);

            dynamic row2 = sheet.ReadRow<ExpandoObjectClass>().Value;
            Assert.Equal("b", row2.Column1);
            Assert.Equal("0", row2.Column2);
            Assert.Equal("0", row2.Column3);
            Assert.Null(row2.Column4);

            dynamic row3 = sheet.ReadRow<ExpandoObjectClass>().Value;
            Assert.Equal("c", row3.Column1);
            Assert.Equal("-2", row3.Column2);
            Assert.Equal("-1", row3.Column3);
            Assert.Null(row3.Column4);
        }

        [Fact]
        public void ReadRow_CustomMappedExpandoObject_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
            importer.Configuration.RegisterClassMap(new CustomExpandoObjectClassMap());

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            dynamic row1 = sheet.ReadRow<ExpandoObjectClass>().Value;
            Assert.Equal("1", row1.Column2);
            Assert.Equal("2", row1.Column3);

            dynamic row2 = sheet.ReadRow<ExpandoObjectClass>().Value;
            Assert.Equal("0", row2.Column2);
            Assert.Equal("0", row2.Column3);

            dynamic row3 = sheet.ReadRow<ExpandoObjectClass>().Value;
            Assert.Equal("-2", row3.Column2);
            Assert.Equal("-1", row3.Column3);
        }

        [Fact]
        public void ReadRow_ExpandoObject_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            dynamic row1 = sheet.ReadRow<ExpandoObject>();
            Assert.Equal("a", row1.Column1);
            Assert.Equal("1", row1.Column2);
            Assert.Equal("2", row1.Column3);
            Assert.Null(row1.Column4);

            dynamic row2 = sheet.ReadRow<ExpandoObject>();
            Assert.Equal("b", row2.Column1);
            Assert.Equal("0", row2.Column2);
            Assert.Equal("0", row2.Column3);
            Assert.Null(row2.Column4);

            dynamic row3 = sheet.ReadRow<ExpandoObject>();
            Assert.Equal("c", row3.Column1);
            Assert.Equal("-2", row3.Column2);
            Assert.Equal("-1", row3.Column3);
            Assert.Null(row3.Column4);
        }

        [Fact]
        public void ReadRow_ExpandoObjectNoHeading_ThrowsExcelMappingException()
        {
            using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.HasHeading = false;

            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<ExpandoObjectClass>());
        }

        private class ExpandoObjectClass
        {
            public ExpandoObject Value { get; set; }
        }

        private class DefaultExpandoObjectClassMap : ExcelClassMap<ExpandoObjectClass>
        {
            public DefaultExpandoObjectClassMap()
            {
                Map(p => p.Value);
            }
        }

        private class CustomExpandoObjectClassMap : ExcelClassMap<ExpandoObjectClass>
        {
            public CustomExpandoObjectClassMap()
            {
                Map(p => p.Value)
                    .WithColumnNames("Column2", "Column3");
            }
        }
    }
}

using System;
using System.Collections.Generic;
using Xunit;

namespace ExcelMapper.Tests
{
    public class MapDictionaryTest
    {
        [Fact]
        public void ReadRow_AutoMappedIDictionaryStringObjectClass_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            IDictionaryStringObjectClass row1 = sheet.ReadRow<IDictionaryStringObjectClass>();
            Assert.Equal(4, row1.Value.Count);
            Assert.Equal("a", row1.Value["Column1"]);
            Assert.Equal("1", row1.Value["Column2"]);
            Assert.Equal("2", row1.Value["Column3"]);
            Assert.Null(row1.Value["Column4"]);

            IDictionaryStringObjectClass row2 = sheet.ReadRow<IDictionaryStringObjectClass>();
            Assert.Equal(4, row2.Value.Count);
            Assert.Equal("b", row2.Value["Column1"]);
            Assert.Equal("0", row2.Value["Column2"]);
            Assert.Equal("0", row2.Value["Column3"]);
            Assert.Null(row2.Value["Column4"]);

            IDictionaryStringObjectClass row3 = sheet.ReadRow<IDictionaryStringObjectClass>();
            Assert.Equal(4, row3.Value.Count);
            Assert.Equal("c", row3.Value["Column1"]);
            Assert.Equal("-2", row3.Value["Column2"]);
            Assert.Equal("-1", row3.Value["Column3"]);
            Assert.Null(row3.Value["Column4"]);
        }

        [Fact]
        public void ReadRow_AutoMappedIDictionaryStringIntClass_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            IDictionaryStringIntClass row1 = sheet.ReadRow<IDictionaryStringIntClass>();
            Assert.Equal(2, row1.Value.Count);
            Assert.Equal(1, row1.Value["Column1"]);
            Assert.Equal(2, row1.Value["Column2"]);

            IDictionaryStringIntClass row2 = sheet.ReadRow<IDictionaryStringIntClass>();
            Assert.Equal(2, row2.Value.Count);
            Assert.Equal(0, row2.Value["Column1"]);
            Assert.Equal(0, row2.Value["Column2"]);

            IDictionaryStringIntClass row3 = sheet.ReadRow<IDictionaryStringIntClass>();
            Assert.Equal(2, row3.Value.Count);
            Assert.Equal(-2, row3.Value["Column1"]);
            Assert.Equal(-1, row3.Value["Column2"]);
        }

        [Fact]
        public void ReadRow_AutoMappedDictionaryStringObject_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            DictionaryStringObjectClass row1 = sheet.ReadRow<DictionaryStringObjectClass>();
            Assert.Equal(4, row1.Value.Count);
            Assert.Equal("a", row1.Value["Column1"]);
            Assert.Equal("1", row1.Value["Column2"]);
            Assert.Equal("2", row1.Value["Column3"]);
            Assert.Null(row1.Value["Column4"]);

            DictionaryStringObjectClass row2 = sheet.ReadRow<DictionaryStringObjectClass>();
            Assert.Equal(4, row2.Value.Count);
            Assert.Equal("b", row2.Value["Column1"]);
            Assert.Equal("0", row2.Value["Column2"]);
            Assert.Equal("0", row2.Value["Column3"]);
            Assert.Null(row2.Value["Column4"]);

            DictionaryStringObjectClass row3 = sheet.ReadRow<DictionaryStringObjectClass>();
            Assert.Equal(4, row3.Value.Count);
            Assert.Equal("c", row3.Value["Column1"]);
            Assert.Equal("-2", row3.Value["Column2"]);
            Assert.Equal("-1", row3.Value["Column3"]);
            Assert.Null(row3.Value["Column4"]);
        }

        [Fact]
        public void ReadRow_AutoMappedDictionaryStringIntClass_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            DictionaryStringIntClass row1 = sheet.ReadRow<DictionaryStringIntClass>();
            Assert.Equal(2, row1.Value.Count);
            Assert.Equal(1, row1.Value["Column1"]);
            Assert.Equal(2, row1.Value["Column2"]);

            DictionaryStringIntClass row2 = sheet.ReadRow<DictionaryStringIntClass>();
            Assert.Equal(2, row2.Value.Count);
            Assert.Equal(0, row2.Value["Column1"]);
            Assert.Equal(0, row2.Value["Column2"]);

            DictionaryStringIntClass row3 = sheet.ReadRow<DictionaryStringIntClass>();
            Assert.Equal(2, row3.Value.Count);
            Assert.Equal(-2, row3.Value["Column1"]);
            Assert.Equal(-1, row3.Value["Column2"]);
        }

        [Fact]
        public void ReadRow_AutoMappedDictionaryStringInvalidObject_ThrowsMissingMethodException()
        {
            using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            Assert.Throws<MissingMethodException>(() => sheet.ReadRow<DictionaryStringInvalidClass>());
        }

        [Fact]
        public void ReadRow_DefaultMappedIDictionaryStringObjectClass_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
            importer.Configuration.RegisterClassMap<DefaultIDictionaryStringObjectClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            IDictionaryStringObjectClass row1 = sheet.ReadRow<IDictionaryStringObjectClass>();
            Assert.Equal(4, row1.Value.Count);
            Assert.Equal("a", row1.Value["Column1"]);
            Assert.Equal("1", row1.Value["Column2"]);
            Assert.Equal("2", row1.Value["Column3"]);
            Assert.Null(row1.Value["Column4"]);

            IDictionaryStringObjectClass row2 = sheet.ReadRow<IDictionaryStringObjectClass>();
            Assert.Equal(4, row2.Value.Count);
            Assert.Equal("b", row2.Value["Column1"]);
            Assert.Equal("0", row2.Value["Column2"]);
            Assert.Equal("0", row2.Value["Column3"]);
            Assert.Null(row2.Value["Column4"]);

            IDictionaryStringObjectClass row3 = sheet.ReadRow<IDictionaryStringObjectClass>();
            Assert.Equal(4, row3.Value.Count);
            Assert.Equal("c", row3.Value["Column1"]);
            Assert.Equal("-2", row3.Value["Column2"]);
            Assert.Equal("-1", row3.Value["Column3"]);
            Assert.Null(row3.Value["Column4"]);
        }

        [Fact]
        public void ReadRow_DefaultMappedIDictionaryStringIntClass_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
            importer.Configuration.RegisterClassMap<DefaultIDictionaryStringIntClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            IDictionaryStringIntClass row1 = sheet.ReadRow<IDictionaryStringIntClass>();
            Assert.Equal(2, row1.Value.Count);
            Assert.Equal(1, row1.Value["Column1"]);
            Assert.Equal(2, row1.Value["Column2"]);

            IDictionaryStringIntClass row2 = sheet.ReadRow<IDictionaryStringIntClass>();
            Assert.Equal(2, row2.Value.Count);
            Assert.Equal(0, row2.Value["Column1"]);
            Assert.Equal(0, row2.Value["Column2"]);

            IDictionaryStringIntClass row3 = sheet.ReadRow<IDictionaryStringIntClass>();
            Assert.Equal(2, row3.Value.Count);
            Assert.Equal(-2, row3.Value["Column1"]);
            Assert.Equal(-1, row3.Value["Column2"]);
        }

        [Fact]
        public void ReadRow_DefaultMappedDictionaryStringObject_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
            importer.Configuration.RegisterClassMap<DefaultDictionaryStringObjectClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            DictionaryStringObjectClass row1 = sheet.ReadRow<DictionaryStringObjectClass>();
            Assert.Equal(4, row1.Value.Count);
            Assert.Equal("a", row1.Value["Column1"]);
            Assert.Equal("1", row1.Value["Column2"]);
            Assert.Equal("2", row1.Value["Column3"]);
            Assert.Null(row1.Value["Column4"]);

            DictionaryStringObjectClass row2 = sheet.ReadRow<DictionaryStringObjectClass>();
            Assert.Equal(4, row2.Value.Count);
            Assert.Equal("b", row2.Value["Column1"]);
            Assert.Equal("0", row2.Value["Column2"]);
            Assert.Equal("0", row2.Value["Column3"]);
            Assert.Null(row2.Value["Column4"]);

            DictionaryStringObjectClass row3 = sheet.ReadRow<DictionaryStringObjectClass>();
            Assert.Equal(4, row3.Value.Count);
            Assert.Equal("c", row3.Value["Column1"]);
            Assert.Equal("-2", row3.Value["Column2"]);
            Assert.Equal("-1", row3.Value["Column3"]);
            Assert.Null(row3.Value["Column4"]);
        }

        [Fact]
        public void ReadRow_DefaultMappedDictionaryStringIntClass_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("DictionaryIntMap.xlsx");
            importer.Configuration.RegisterClassMap<DefaultDictionaryStringIntClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            DictionaryStringIntClass row1 = sheet.ReadRow<DictionaryStringIntClass>();
            Assert.Equal(2, row1.Value.Count);
            Assert.Equal(1, row1.Value["Column1"]);
            Assert.Equal(2, row1.Value["Column2"]);

            DictionaryStringIntClass row2 = sheet.ReadRow<DictionaryStringIntClass>();
            Assert.Equal(2, row2.Value.Count);
            Assert.Equal(0, row2.Value["Column1"]);
            Assert.Equal(0, row2.Value["Column2"]);

            DictionaryStringIntClass row3 = sheet.ReadRow<DictionaryStringIntClass>();
            Assert.Equal(2, row3.Value.Count);
            Assert.Equal(-2, row3.Value["Column1"]);
            Assert.Equal(-1, row3.Value["Column2"]);
        }

        [Fact]
        public void ReadRow_CustomMappedIDictionaryStringObject_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
            importer.Configuration.RegisterClassMap(new CustomIDictionaryStringObjectClassMap());

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            IDictionaryStringObjectClass row1 = sheet.ReadRow<IDictionaryStringObjectClass>();
            Assert.Equal(2, row1.Value.Count);
            Assert.Equal("1", row1.Value["Column2"]);
            Assert.Equal("2", row1.Value["Column3"]);

            IDictionaryStringObjectClass row2 = sheet.ReadRow<IDictionaryStringObjectClass>();
            Assert.Equal(2, row2.Value.Count);
            Assert.Equal("0", row2.Value["Column2"]);
            Assert.Equal("0", row2.Value["Column3"]);

            IDictionaryStringObjectClass row3 = sheet.ReadRow<IDictionaryStringObjectClass>();
            Assert.Equal(2, row3.Value.Count);
            Assert.Equal("-2", row3.Value["Column2"]);
            Assert.Equal("-1", row3.Value["Column3"]);
        }

        [Fact]
        public void ReadRow_CustomMappedIDictionaryStringInt_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
            importer.Configuration.RegisterClassMap(new CustomIDictionaryStringIntClassMap());

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            IDictionaryStringIntClass row1 = sheet.ReadRow<IDictionaryStringIntClass>();
            Assert.Equal(2, row1.Value.Count);
            Assert.Equal(1, row1.Value["Column2"]);
            Assert.Equal(2, row1.Value["Column3"]);

            IDictionaryStringIntClass row2 = sheet.ReadRow<IDictionaryStringIntClass>();
            Assert.Equal(2, row2.Value.Count);
            Assert.Equal(0, row2.Value["Column2"]);
            Assert.Equal(0, row2.Value["Column3"]);

            IDictionaryStringIntClass row3 = sheet.ReadRow<IDictionaryStringIntClass>();
            Assert.Equal(2, row3.Value.Count);
            Assert.Equal(-2, row3.Value["Column2"]);
            Assert.Equal(-1, row3.Value["Column3"]);
        }

        [Fact]
        public void ReadRow_CustomMappedDictionaryStringObject_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
            importer.Configuration.RegisterClassMap(new CustomDictionaryStringObjectClassMap());

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            DictionaryStringObjectClass row1 = sheet.ReadRow<DictionaryStringObjectClass>();
            Assert.Equal(2, row1.Value.Count);
            Assert.Equal("1", row1.Value["Column2"]);
            Assert.Equal("2", row1.Value["Column3"]);

            DictionaryStringObjectClass row2 = sheet.ReadRow<DictionaryStringObjectClass>();
            Assert.Equal(2, row2.Value.Count);
            Assert.Equal("0", row2.Value["Column2"]);
            Assert.Equal("0", row2.Value["Column3"]);

            DictionaryStringObjectClass row3 = sheet.ReadRow<DictionaryStringObjectClass>();
            Assert.Equal(2, row3.Value.Count);
            Assert.Equal("-2", row3.Value["Column2"]);
            Assert.Equal("-1", row3.Value["Column3"]);
        }

        [Fact]
        public void ReadRow_CustomMappedDictionaryStringInt_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
            importer.Configuration.RegisterClassMap(new CustomDictionaryStringIntClassMap());

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            DictionaryStringIntClass row1 = sheet.ReadRow<DictionaryStringIntClass>();
            Assert.Equal(2, row1.Value.Count);
            Assert.Equal(1, row1.Value["Column2"]);
            Assert.Equal(2, row1.Value["Column3"]);

            DictionaryStringIntClass row2 = sheet.ReadRow<DictionaryStringIntClass>();
            Assert.Equal(2, row2.Value.Count);
            Assert.Equal(0, row2.Value["Column2"]);
            Assert.Equal(0, row2.Value["Column3"]);

            DictionaryStringIntClass row3 = sheet.ReadRow<DictionaryStringIntClass>();
            Assert.Equal(2, row3.Value.Count);
            Assert.Equal(-2, row3.Value["Column2"]);
            Assert.Equal(-1, row3.Value["Column3"]);
        }

        [Fact]
        public void ReadRow_CustomMapSortedDictionaryStringInt_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
            importer.Configuration.RegisterClassMap(new CustomSortedDictionaryStringIntClassMap());

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            SortedDictionaryStringIntClass row1 = sheet.ReadRow<SortedDictionaryStringIntClass>();
            Assert.Equal(2, row1.Value.Count);
            Assert.Equal(1, row1.Value["Column2"]);
            Assert.Equal(2, row1.Value["Column3"]);

            SortedDictionaryStringIntClass row2 = sheet.ReadRow<SortedDictionaryStringIntClass>();
            Assert.Equal(2, row2.Value.Count);
            Assert.Equal(0, row2.Value["Column2"]);
            Assert.Equal(0, row2.Value["Column3"]);

            SortedDictionaryStringIntClass row3 = sheet.ReadRow<SortedDictionaryStringIntClass>();
            Assert.Equal(2, row3.Value.Count);
            Assert.Equal(-2, row3.Value["Column2"]);
            Assert.Equal(-1, row3.Value["Column3"]);
        }

#if false
        [Fact]
        public void ReadRow_DictionaryObject_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            Dictionary<string, object> row1 = sheet.ReadRow<Dictionary<string, object>>();
            Assert.Equal(4, row1.Count);
            Assert.Equal("a", row1["Column1"]);
            Assert.Equal("1", row1["Column2"]);
            Assert.Equal("2", row1["Column3"]);
            Assert.Null(row1["Column4"]);

            IDictionary<string, string> row2 = sheet.ReadRow<IDictionary<string, string>>();
            Assert.Equal(4, row2.Count);
            Assert.Equal("b", row2["Column1"]);
            Assert.Equal("0", row2["Column2"]);
            Assert.Equal("0", row2["Column3"]);
            Assert.Null(row2["Column4"]);

            IDictionaryStringObjectClass row3 = sheet.ReadRow<IDictionaryStringObjectClass>();
            Assert.Equal(4, row3.Value.Count);
            Assert.Equal("c", row3.Value["Column1"]);
            Assert.Equal("-2", row3.Value["Column2"]);
            Assert.Equal("-1", row3.Value["Column3"]);
            Assert.Null(row3.Value["Column4"]);
        }
#endif

        [Fact]
        public void ReadRow_IDictionaryStringObject_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IDictionary<string, object>>());
        }

        [Fact]
        public void ReadRow_DictionaryStringObject_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Dictionary<string, object>>());
        }

        [Fact]
        public void ReadRow_DictionaryNoHeading_ThrowsExcelMappingException()
        {
            using var importer = Helpers.GetImporter("DictionaryMap.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.HasHeading = false;

            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DictionaryStringObjectClass>());
        }

        [Fact]
        public void ReadRow_DictionaryNoHeadingWithCustomMap_ThrowsExcelMappingException()
        {
            using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
            importer.Configuration.RegisterClassMap<DefaultDictionaryStringObjectClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.HasHeading = false;

            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DictionaryStringObjectClass>());
        }

        private class IDictionaryStringObjectClass
        {
            public IDictionary<string, object> Value { get; set; }
        }

        private class DefaultIDictionaryStringObjectClassMap : ExcelClassMap<IDictionaryStringObjectClass>
        {
            public DefaultIDictionaryStringObjectClassMap()
            {
                Map(p => p.Value);
            }
        }

        private class CustomIDictionaryStringObjectClassMap : ExcelClassMap<IDictionaryStringObjectClass>
        {
            public CustomIDictionaryStringObjectClassMap()
            {
                Map(p => p.Value)
                    .WithColumnNames("Column2", "Column3");
            }
        }

        private class IDictionaryStringIntClass
        {
            public IDictionary<string, int> Value { get; set; }
        }

        private class DefaultIDictionaryStringIntClassMap : ExcelClassMap<IDictionaryStringIntClass>
        {
            public DefaultIDictionaryStringIntClassMap()
            {
                Map(p => p.Value);
            }
        }

        private class CustomIDictionaryStringIntClassMap : ExcelClassMap<IDictionaryStringIntClass>
        {
            public CustomIDictionaryStringIntClassMap()
            {
                Map(p => p.Value)
                    .WithColumnNames("Column2", "Column3");
            }
        }

        private class DictionaryStringObjectClass
        {
            public Dictionary<string, object> Value { get; set; }
        }

        private class DefaultDictionaryStringObjectClassMap : ExcelClassMap<DictionaryStringObjectClass>
        {
            public DefaultDictionaryStringObjectClassMap()
            {
                Map(p => p.Value);
            }
        }

        private class CustomDictionaryStringObjectClassMap : ExcelClassMap<DictionaryStringObjectClass>
        {
            public CustomDictionaryStringObjectClassMap()
            {
                Map(p => p.Value)
                    .WithColumnNames("Column2", "Column3");
            }
        }

        private class DictionaryStringIntClass
        {
            public Dictionary<string, int> Value { get; set; }
        }

        private class DefaultDictionaryStringIntClassMap : ExcelClassMap<DictionaryStringIntClass>
        {
            public DefaultDictionaryStringIntClassMap()
            {
                Map(p => p.Value);
            }
        }

        private class CustomDictionaryStringIntClassMap : ExcelClassMap<DictionaryStringIntClass>
        {
            public CustomDictionaryStringIntClassMap()
            {
                Map(p => p.Value)
                    .WithColumnNames("Column2", "Column3");
            }
        }

        private class DictionaryStringInvalidClass
        {
            public Dictionary<string, ExcelSheet> Value { get; set; }
        }

        private class SortedDictionaryStringIntClass
        {
            public SortedDictionary<string, int> Value { get; set; }
        }

        private class CustomSortedDictionaryStringIntClassMap : ExcelClassMap<SortedDictionaryStringIntClass>
        {
            public CustomSortedDictionaryStringIntClassMap()
            {
                Map<int>(p => p.Value)
                    .WithColumnNames("Column2", "Column3");
            }
        }

        private class DefaultSortedDictionaryStringIntClassMap : ExcelClassMap<SortedDictionaryStringIntClass>
        {
            public DefaultSortedDictionaryStringIntClassMap()
            {
                Map<int>(p => p.Value)
                    .WithColumnNames("Column2", "Column3");
            }
        }

        [Fact]
        public void ReadRow_DictionaryMissingColumn_ThrowsExcelMappingException()
        {
            using var importer = Helpers.GetImporter("DictionaryMap.xlsx");
            importer.Configuration.RegisterClassMap<MissingColumnClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IDictionaryStringObjectClass>());
        }

        private class MissingColumnClassMap : ExcelClassMap<IDictionaryStringObjectClass>
        {
            public MissingColumnClassMap()
            {
                Map(p => p.Value)
                    .WithColumnNames("NoSuchColumn");
            }
        }
    }
}

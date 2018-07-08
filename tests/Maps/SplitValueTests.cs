using System.Collections.Generic;
using Xunit;

namespace ExcelMapper.Tests
{
    public class SplitValueTests
    {
        [Fact]
        public void ReadRow_AutoMappedEnumerable_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("SplitWithComma.xlsx"))
            {
                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                AutoSplitWithSeparatorClass row1 = sheet.ReadRow<AutoSplitWithSeparatorClass>();
                Assert.Equal(new string[] { "1", "2", "3" }, row1.Value);

                AutoSplitWithSeparatorClass row2 = sheet.ReadRow<AutoSplitWithSeparatorClass>();
                Assert.Equal(new string[] { "1", null, "2" }, row2.Value);
            }
        }

        [Fact]
        public void ReadRow_CustomElementMap_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("SplitWithComma.xlsx"))
            {
                importer.Configuration.RegisterClassMap<SplitWithElementMapMap>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                AutoSplitWithSeparatorClass row1 = sheet.ReadRow<AutoSplitWithSeparatorClass>();
                Assert.Equal(new string[] { "1", "2", "3" }, row1.Value);

                AutoSplitWithSeparatorClass row2 = sheet.ReadRow<AutoSplitWithSeparatorClass>();
                Assert.Equal(new string[] { "1", "empty", "2" }, row2.Value);
            }
        }

        [Fact]
        public void ReadRow_SeparatorsArrayMap_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("SplitWithCustomSeparators.xlsx"))
            {
                importer.Configuration.RegisterClassMap<SplitWithSeparatorsArrayMap>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                AutoSplitWithSeparatorClass row1 = sheet.ReadRow<AutoSplitWithSeparatorClass>();
                Assert.Equal(new string[] { "1", "2", "3", "4", "5" }, row1.Value);

                AutoSplitWithSeparatorClass row2 = sheet.ReadRow<AutoSplitWithSeparatorClass>();
                Assert.Equal(new string[] { "1", "2", "3" }, row2.Value);
            }
        }

        [Fact]
        public void ReadRow_IEnumerableSeparatorsMap_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("SplitWithCustomSeparators.xlsx"))
            {
                importer.Configuration.RegisterClassMap<SplitWithEnumerableSeparatorsMap>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                AutoSplitWithSeparatorClass row1 = sheet.ReadRow<AutoSplitWithSeparatorClass>();
                Assert.Equal(new string[] { "1", "2", "3", "4", "5" }, row1.Value);

                AutoSplitWithSeparatorClass row2 = sheet.ReadRow<AutoSplitWithSeparatorClass>();
                Assert.Equal(new string[] { "1", "2", "3" }, row2.Value);
            }
        }

        [Fact]
        public void ReadRow_CustomColumnName_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("SplitWithComma.xlsx"))
            {
                importer.Configuration.RegisterClassMap(new SplitWithSeparatorMap());

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                CustomSplitWithSeparatorClass row1 = sheet.ReadRow<CustomSplitWithSeparatorClass>();
                Assert.Equal(new string[] { "1", "2", "3" }, row1.Value);

                Assert.Equal(new string[] { "1", "2", "3" }, row1.ValueWithColumnName);
                Assert.Equal(new string[] { "1", "2", "3" }, row1.ValueWithColumnIndex);

                Assert.Equal(new string[] { "1", "2", "3" }, row1.ValueWithColumnNameAcrossMultiColumnNames);
                Assert.Equal(new string[] { "1", "2", "3" }, row1.ValueWithColumnNameAcrossMultiColumnIndices);

                Assert.Equal(new string[] { "1", "2", "3" }, row1.ValueWithColumnIndexAcrossMultiColumnNames);
                Assert.Equal(new string[] { "1", "2", "3" }, row1.ValueWithColumnIndexAcrossMultiColumnIndices);
            }
        }

        [Fact]
        public void ReadRow_NullableIntArray_ReturnsExpected()
        {

        }

        private class AutoSplitWithSeparatorClass
        {
            public string[] Value { get; set; }
        }

        private class CustomSplitWithSeparatorClass
        {
            public string[] Value { get; set; }
            public string[] ValueWithColumnName { get; set; }
            public string[] ValueWithColumnIndex { get; set; }

            public string[] ValueWithColumnNameAcrossMultiColumnNames { get; set; }
            public string[] ValueWithColumnNameAcrossMultiColumnIndices { get; set; }

            public string[] ValueWithColumnIndexAcrossMultiColumnNames { get; set; }
            public string[] ValueWithColumnIndexAcrossMultiColumnIndices { get; set; }
        }

        private class SplitWithElementMapMap : ExcelClassMap<AutoSplitWithSeparatorClass>
        {
            public SplitWithElementMapMap()
            {
                Map(p => p.Value)
                    .WithElementMap(e => e
                        .WithEmptyFallback("empty")
                    );
            }
        }

        private class SplitWithSeparatorsArrayMap : ExcelClassMap<AutoSplitWithSeparatorClass>
        {
            public SplitWithSeparatorsArrayMap()
            {
                Map(p => p.Value)
                    .WithSeparators(';', ',');
            }
        }

        private class SplitWithEnumerableSeparatorsMap : ExcelClassMap<AutoSplitWithSeparatorClass>
        {
            public SplitWithEnumerableSeparatorsMap()
            {
                Map(p => p.Value)
                    .WithSeparators(new List<char> {';', ','});
            }
        }

        private class SplitWithSeparatorMap : ExcelClassMap<CustomSplitWithSeparatorClass>
        {
            public SplitWithSeparatorMap()
            {
                Map(p => p.Value);

                Map(p => p.ValueWithColumnName)
                    .WithColumnName("Value");

                Map(p => p.ValueWithColumnIndex)
                    .WithColumnIndex(0);

                Map(p => p.ValueWithColumnNameAcrossMultiColumnNames)
                    .WithColumnNames("IListString1", "IListString2")
                    .WithColumnName("Value");

                Map(p => p.ValueWithColumnNameAcrossMultiColumnIndices)
                    .WithColumnIndices(9, 10)
                    .WithColumnName("Value");

                Map(p => p.ValueWithColumnIndexAcrossMultiColumnNames)
                    .WithColumnNames("IListString1", "IListString2")
                    .WithColumnIndex(0);

                Map(p => p.ValueWithColumnIndexAcrossMultiColumnIndices)
                    .WithColumnIndices(9, 10)
                    .WithColumnIndex(0);
            }
        }
    }
}

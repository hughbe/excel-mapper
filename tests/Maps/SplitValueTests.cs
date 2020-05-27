using System.Collections.Generic;
using System.Collections.ObjectModel;
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
                Assert.Equal(new ObservableCollectionEnum[] { ObservableCollectionEnum.Value1, ObservableCollectionEnum.Value2, ObservableCollectionEnum.Value3 }, row1.EnumValue);

                // Empty value.
                Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<AutoSplitWithSeparatorClass>());

                // Invalid value.
                AutoSplitWithSeparatorClass row3 = sheet.ReadRow<AutoSplitWithSeparatorClass>();
                Assert.Equal(new string[] { "1" }, row3.Value);

                Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<AutoSplitWithSeparatorClass>());
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

                MappedAutoSplitWithSeparatorClass row1 = sheet.ReadRow<MappedAutoSplitWithSeparatorClass>();
                Assert.Equal(new string[] { "1", "2", "3" }, row1.Value);
                Assert.Equal(new ObservableCollectionEnum[] { ObservableCollectionEnum.Value1, ObservableCollectionEnum.Value2, ObservableCollectionEnum.Value3 }, row1.EnumCollectionValue);
                Assert.Equal(new ObservableCollectionEnum[] { ObservableCollectionEnum.Value1, ObservableCollectionEnum.Value2, ObservableCollectionEnum.Value3 }, row1.EnumValue);

                MappedAutoSplitWithSeparatorClass row2 = sheet.ReadRow<MappedAutoSplitWithSeparatorClass>();
                Assert.Equal(new string[] { "1", "empty", "2" }, row2.Value);
                Assert.Equal(new ObservableCollectionEnum[] { ObservableCollectionEnum.Value1, ObservableCollectionEnum.Empty, ObservableCollectionEnum.Value3 }, row2.EnumCollectionValue);
                Assert.Equal(new ObservableCollectionEnum[] { ObservableCollectionEnum.Value1, ObservableCollectionEnum.Empty, ObservableCollectionEnum.Value3 }, row2.EnumValue);

                MappedAutoSplitWithSeparatorClass row3 = sheet.ReadRow<MappedAutoSplitWithSeparatorClass>();
                Assert.Equal(new string[] { "1" }, row3.Value);
                Assert.Equal(new ObservableCollectionEnum[] { ObservableCollectionEnum.Value1 }, row3.EnumCollectionValue);
                Assert.Equal(new ObservableCollectionEnum[] { ObservableCollectionEnum.Value1 }, row3.EnumValue);

                MappedAutoSplitWithSeparatorClass row4 = sheet.ReadRow<MappedAutoSplitWithSeparatorClass>();
                Assert.Equal(new string[0], row4.Value);
                Assert.Equal(new ObservableCollectionEnum[] { ObservableCollectionEnum.Invalid }, row4.EnumCollectionValue);
                Assert.Equal(new ObservableCollectionEnum[] { ObservableCollectionEnum.Invalid }, row4.EnumValue);
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

        private class AutoSplitWithSeparatorClass
        {
            public string[] Value { get; set; }
            public ObservableCollection<ObservableCollectionEnum> EnumValue { get; set; }
        }

        private class MappedAutoSplitWithSeparatorClass
        {
            public string[] Value { get; set; }
            public Collection<ObservableCollectionEnum> EnumCollectionValue { get; set; }
            public ObservableCollection<ObservableCollectionEnum> EnumValue { get; set; }
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

        private class SplitWithElementMapMap : ExcelClassMap<MappedAutoSplitWithSeparatorClass>
        {
            public SplitWithElementMapMap()
            {
                Map(p => p.Value)
                    .WithElementMap(e => e
                        .WithEmptyFallback("empty")
                    );

                Map(p => p.EnumCollectionValue)
                    .WithColumnName("EnumValue")
                    .WithElementMap(e => e
                        .WithEmptyFallback(ObservableCollectionEnum.Empty)
                        .WithInvalidFallback(ObservableCollectionEnum.Invalid)
                    );

                Map(p => p.EnumValue)
                    .WithElementMap(e => e
                        .WithEmptyFallback(ObservableCollectionEnum.Empty)
                        .WithInvalidFallback(ObservableCollectionEnum.Invalid)
                    );
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

        public enum ObservableCollectionEnum
        {
            Value1,
            Value2,
            Value3,
            Empty,
            Invalid
        }
    }
}

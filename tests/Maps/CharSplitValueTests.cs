using System.Collections.Generic;
using System.Collections.ObjectModel;
using Xunit;

namespace ExcelMapper.Tests
{
    public class CharSplitValueTests
    {
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

        private class AutoSplitWithSeparatorClass
        {
            public string[] Value { get; set; }
            public ObservableCollection<ObservableCollectionEnum> EnumValue { get; set; }
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
                    .WithSeparators(new List<char> { ';', ',' });
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

        [Fact]
        public void ReadRow_MultiMapMissingRow_ThrowsExcelMappingException()
        {
            using var importer = Helpers.GetImporter("SplitWithCustomSeparators.xlsx");
            importer.Configuration.RegisterClassMap<MissingColumnRowMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<MissingColumnRow>());
        }

        [Fact]
        public void ReadRow_MultiMapOptionalMissingRow_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("SplitWithCustomSeparators.xlsx");
            importer.Configuration.RegisterClassMap<OptionalMissingColumnRowMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            MissingColumnRow row = sheet.ReadRow<MissingColumnRow>();
            Assert.Null(row.MissingValue);
        }

        private class MissingColumnRow
        {
            public int[] MissingValue { get; set; }
        }

        private class MissingColumnRowMap : ExcelClassMap<MissingColumnRow>
        {
            public MissingColumnRowMap()
            {
                Map(p => p.MissingValue)
                    .WithSeparators(';', ',');
            }
        }

        private class OptionalMissingColumnRowMap : ExcelClassMap<MissingColumnRow>
        {
            public OptionalMissingColumnRowMap()
            {
                Map(p => p.MissingValue)
                    .WithSeparators(';', ',')
                    .MakeOptional();
            }
        }
    }
}

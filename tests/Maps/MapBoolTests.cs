using Xunit;

namespace ExcelMapper.Tests
{
    public class MapBoolTests
    {
        [Fact]
        public void ReadRow_AutoMappedBool_Success()
        {
            using (var importer = Helpers.GetImporter("Bools.xlsx"))
            {
                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                // Valid cell value.
                BoolValue row1 = sheet.ReadRow<BoolValue>();
                Assert.True(row1.Value);

                BoolValue row2 = sheet.ReadRow<BoolValue>();
                Assert.True(row2.Value);

                BoolValue row3 = sheet.ReadRow<BoolValue>();
                Assert.False(row3.Value);

                BoolValue row4 = sheet.ReadRow<BoolValue>();
                Assert.False(row4.Value);

                // Empty cell value.
                Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<BoolValue>());

                // Invalid cell value.
                Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<BoolValue>());
            }
        }

        [Fact]
        public void ReadRow_AutoMappedNullableBool_Success()
        {
            using (var importer = Helpers.GetImporter("Bools.xlsx"))
            {
                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                // Valid cell value.
                NullableBoolValue row1 = sheet.ReadRow<NullableBoolValue>();
                Assert.True(row1.Value);

                NullableBoolValue row2 = sheet.ReadRow<NullableBoolValue>();
                Assert.True(row2.Value);

                NullableBoolValue row3 = sheet.ReadRow<NullableBoolValue>();
                Assert.False(row3.Value);

                NullableBoolValue row4 = sheet.ReadRow<NullableBoolValue>();
                Assert.False(row4.Value);

                // Empty cell value.
                NullableBoolValue row5 = sheet.ReadRow<NullableBoolValue>();
                Assert.Null(row5.Value);

                // Invalid cell value.
                Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableBoolValue>());
            }
        }

        [Fact]
        public void ReadRow_CustomMappedBool_Success()
        {
            using (var importer = Helpers.GetImporter("Bools.xlsx"))
            {
                importer.Configuration.RegisterClassMap<BoolValueFallbackMap>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                // Valid cell value.
                BoolValue row1 = sheet.ReadRow<BoolValue>();
                Assert.True(row1.Value);

                BoolValue row2 = sheet.ReadRow<BoolValue>();
                Assert.True(row2.Value);

                BoolValue row3 = sheet.ReadRow<BoolValue>();
                Assert.False(row3.Value);

                BoolValue row4 = sheet.ReadRow<BoolValue>();
                Assert.False(row4.Value);

                // Empty cell value.
                BoolValue row5 = sheet.ReadRow<BoolValue>();
                Assert.True(row5.Value);

                // Invalid cell value.
                BoolValue row6 = sheet.ReadRow<BoolValue>();
                Assert.True(row6.Value);
            }
        }

        [Fact]
        public void ReadRow_CustomMappedNullableBool_Success()
        {
            using (var importer = Helpers.GetImporter("Bools.xlsx"))
            {
                importer.Configuration.RegisterClassMap<NullableBoolValueFallbackMap>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                // Valid cell value.
                NullableBoolValue row1 = sheet.ReadRow<NullableBoolValue>();
                Assert.True(row1.Value);

                NullableBoolValue row2 = sheet.ReadRow<NullableBoolValue>();
                Assert.True(row2.Value);

                NullableBoolValue row3 = sheet.ReadRow<NullableBoolValue>();
                Assert.False(row3.Value);

                NullableBoolValue row4 = sheet.ReadRow<NullableBoolValue>();
                Assert.False(row4.Value);

                // Empty cell value.
                NullableBoolValue row5 = sheet.ReadRow<NullableBoolValue>();
                Assert.True(row5.Value);

                // Invalid cell value.
                NullableBoolValue row6 = sheet.ReadRow<NullableBoolValue>();
                Assert.True(row6.Value);
            }
        }

        private class BoolValue
        {
            public bool Value { get; set; }
        }

        private class NullableBoolValue
        {
            public bool? Value { get; set; }
        }

        private class BoolValueFallbackMap : ExcelClassMap<BoolValue>
        {
            public BoolValueFallbackMap()
            {
                Map(o => o.Value)
                    .WithEmptyFallback(true)
                    .WithInvalidFallback(true);
            }
        }

        private class NullableBoolValueFallbackMap : ExcelClassMap<NullableBoolValue>
        {
            public NullableBoolValueFallbackMap()
            {
                Map(o => o.Value)
                    .WithEmptyFallback(true)
                    .WithInvalidFallback(true);
            }
        }
    }
}

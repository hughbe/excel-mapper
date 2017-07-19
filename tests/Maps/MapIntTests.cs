using Xunit;

namespace ExcelMapper.Tests
{
    public class MapIntests
    {
        [Fact]
        public void ReadRow_AutoMappedInt_Success()
        {
            using (var importer = Helpers.GetImporter("Numbers.xlsx"))
            {
                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                // Valid cell value.
                IntValue row1 = sheet.ReadRow<IntValue>();
                Assert.Equal(2, row1.Value);

                // Empty cell value.
                Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IntValue>());

                // Invalid cell value.
                Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IntValue>());
            }
        }

        [Fact]
        public void ReadRow_AutoMappedNullableInt_Success()
        {
            using (var importer = Helpers.GetImporter("Numbers.xlsx"))
            {
                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                // Valid cell value.
                NullableIntValue row1 = sheet.ReadRow<NullableIntValue>();
                Assert.Equal(2, row1.Value);

                // Empty cell value.
                NullableIntValue row2 = sheet.ReadRow<NullableIntValue>();
                Assert.Null(row2.Value);

                // Invalid cell value.
                Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<IntValue>());
            }
        }

        [Fact]
        public void ReadRow_CustomMappedInt_Success()
        {
            using (var importer = Helpers.GetImporter("Numbers.xlsx"))
            {
                importer.Configuration.RegisterClassMap<IntValueFallbackMap>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                // Valid cell value.
                IntValue row1 = sheet.ReadRow<IntValue>();
                Assert.Equal(2, row1.Value);

                // Empty cell value.
                IntValue row2 = sheet.ReadRow<IntValue>();
                Assert.Equal(-10, row2.Value);

                // Invalid cell value.
                IntValue row3 = sheet.ReadRow<IntValue>();
                Assert.Equal(10, row3.Value);
            }
        }

        [Fact]
        public void ReadRow_CustomMappedNullableInt_Success()
        {
            using (var importer = Helpers.GetImporter("Numbers.xlsx"))
            {
                importer.Configuration.RegisterClassMap<NullableIntValueFallbackMap>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                // Valid cell value.
                NullableIntValue row1 = sheet.ReadRow<NullableIntValue>();
                Assert.Equal(2, row1.Value);

                // Empty cell value.
                NullableIntValue row2 = sheet.ReadRow<NullableIntValue>();
                Assert.Equal(-10, row2.Value);

                // Invalid cell value.
                NullableIntValue row3 = sheet.ReadRow<NullableIntValue>();
                Assert.Equal(10, row3.Value);
            }
        }

        private class IntValue
        {
            public int Value { get; set; }
        }

        private class NullableIntValue
        {
            public int? Value { get; set; }
        }

        private class IntValueFallbackMap : ExcelClassMap<IntValue>
        {
            public IntValueFallbackMap()
            {
                Map(o => o.Value)
                    .WithEmptyFallback(-10)
                    .WithInvalidFallback(10);
            }
        }

        private class NullableIntValueFallbackMap : ExcelClassMap<NullableIntValue>
        {
            public NullableIntValueFallbackMap()
            {
                Map(o => o.Value)
                    .WithEmptyFallback(-10)
                    .WithInvalidFallback(10);
            }
        }
    }
}

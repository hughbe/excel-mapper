using Xunit;

namespace ExcelMapper.Tests
{
    public class MapDecimalTests
    {
        [Fact]
        public void ReadRow_AutoMappedDecimal_Success()
        {
            using (var importer = Helpers.GetImporter("Doubles.xlsx"))
            {
                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                // Valid cell value.
                DecimalValue row1 = sheet.ReadRow<DecimalValue>();
                Assert.Equal(2.2345m, row1.Value);

                // Empty cell value.
                Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DecimalValue>());

                // Invalid cell value.
                Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DecimalValue>());
            }
        }

        [Fact]
        public void ReadRow_AutoMappedNullableDecimal_Success()
        {
            using (var importer = Helpers.GetImporter("Doubles.xlsx"))
            {
                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                // Valid cell value.
                NullableDecimalValue row1 = sheet.ReadRow<NullableDecimalValue>();
                Assert.Equal(2.2345m, row1.Value);

                // Empty cell value.
                NullableDecimalValue row2 = sheet.ReadRow<NullableDecimalValue>();
                Assert.Null(row2.Value);

                // Invalid cell value.
                Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DecimalValue>());
            }
        }

        [Fact]
        public void ReadRow_CustomMappedDecimal_Success()
        {
            using (var importer = Helpers.GetImporter("Doubles.xlsx"))
            {
                importer.Configuration.RegisterClassMap<DecimalValueFallbackMap>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                // Valid cell value.
                DecimalValue row1 = sheet.ReadRow<DecimalValue>();
                Assert.Equal(2.2345m, row1.Value);

                // Empty cell value.
                DecimalValue row2 = sheet.ReadRow<DecimalValue>();
                Assert.Equal(-10, row2.Value);

                // Invalid cell value.
                DecimalValue row3 = sheet.ReadRow<DecimalValue>();
                Assert.Equal(10, row3.Value);
            }
        }

        [Fact]
        public void ReadRow_CustomMappedNullableDecimal_Success()
        {
            using (var importer = Helpers.GetImporter("Doubles.xlsx"))
            {
                importer.Configuration.RegisterClassMap<NullableDecimalValueFallbackMap>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                // Valid cell value.
                NullableDecimalValue row1 = sheet.ReadRow<NullableDecimalValue>();
                Assert.Equal(2.2345m, row1.Value);

                // Empty cell value.
                NullableDecimalValue row2 = sheet.ReadRow<NullableDecimalValue>();
                Assert.Equal(-10, row2.Value);

                // Invalid cell value.
                NullableDecimalValue row3 = sheet.ReadRow<NullableDecimalValue>();
                Assert.Equal(10, row3.Value);
            }
        }

        private class DecimalValue
        {
            public decimal Value { get; set; }
        }

        private class NullableDecimalValue
        {
            public decimal? Value { get; set; }
        }

        private class DecimalValueFallbackMap : ExcelClassMap<DecimalValue>
        {
            public DecimalValueFallbackMap()
            {
                Map(o => o.Value)
                    .WithEmptyFallback(-10m)
                    .WithInvalidFallback(10m);
            }
        }

        private class NullableDecimalValueFallbackMap : ExcelClassMap<NullableDecimalValue>
        {
            public NullableDecimalValueFallbackMap()
            {
                Map(o => o.Value)
                    .WithEmptyFallback(-10m)
                    .WithInvalidFallback(10m);
            }
        }
    }
}

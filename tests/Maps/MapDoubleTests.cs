using Xunit;

namespace ExcelMapper.Tests
{
    public class MapDoubleTests
    {
        [Fact]
        public void ReadRow_AutoMappedDouble_Success()
        {
            using (var importer = Helpers.GetImporter("Doubles.xlsx"))
            {
                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                // Valid cell value.
                DoubleValue row1 = sheet.ReadRow<DoubleValue>();
                Assert.Equal(2.2345, row1.Value);

                // Empty cell value.
                Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DoubleValue>());

                // Invalid cell value.
                Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DoubleValue>());
            }
        }

        [Fact]
        public void ReadRow_AutoMappedNullableDouble_Success()
        {
            using (var importer = Helpers.GetImporter("Doubles.xlsx"))
            {
                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                // Valid cell value.
                NullableDoubleValue row1 = sheet.ReadRow<NullableDoubleValue>();
                Assert.Equal(2.2345, row1.Value);

                // Empty cell value.
                NullableDoubleValue row2 = sheet.ReadRow<NullableDoubleValue>();
                Assert.Null(row2.Value);

                // Invalid cell value.
                Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableDoubleValue>());
            }
        }

        [Fact]
        public void ReadRow_CustomMappedDouble_Success()
        {
            using (var importer = Helpers.GetImporter("Doubles.xlsx"))
            {
                importer.Configuration.RegisterClassMap<DoubleValueFallbackMap>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                // Valid cell value.
                DoubleValue row1 = sheet.ReadRow<DoubleValue>();
                Assert.Equal(2.2345, row1.Value);

                // Empty cell value.
                DoubleValue row2 = sheet.ReadRow<DoubleValue>();
                Assert.Equal(-10, row2.Value);

                // Invalid cell value.
                DoubleValue row3 = sheet.ReadRow<DoubleValue>();
                Assert.Equal(10, row3.Value);
            }
        }

        [Fact]
        public void ReadRow_CustomMappedNullableDouble_Success()
        {
            using (var importer = Helpers.GetImporter("Doubles.xlsx"))
            {
                importer.Configuration.RegisterClassMap<NullableDoubleValueFallbackMap>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                // Valid cell value.
                NullableDoubleValue row1 = sheet.ReadRow<NullableDoubleValue>();
                Assert.Equal(2.2345, row1.Value);

                // Empty cell value.
                NullableDoubleValue row2 = sheet.ReadRow<NullableDoubleValue>();
                Assert.Equal(-10, row2.Value);

                // Invalid cell value.
                NullableDoubleValue row3 = sheet.ReadRow<NullableDoubleValue>();
                Assert.Equal(10, row3.Value);
            }
        }

        private class DoubleValue
        {
            public double Value { get; set; }
        }

        private class NullableDoubleValue
        {
            public double? Value { get; set; }
        }

        private class DoubleValueFallbackMap : ExcelClassMap<DoubleValue>
        {
            public DoubleValueFallbackMap()
            {
                Map(o => o.Value)
                    .WithEmptyFallback(-10.0)
                    .WithInvalidFallback(10.0);
            }
        }

        private class NullableDoubleValueFallbackMap : ExcelClassMap<NullableDoubleValue>
        {
            public NullableDoubleValueFallbackMap()
            {
                Map(o => o.Value)
                    .WithEmptyFallback(-10.0)
                    .WithInvalidFallback(10.0);
            }
        }
    }
}

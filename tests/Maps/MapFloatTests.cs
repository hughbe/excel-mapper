using Xunit;

namespace ExcelMapper.Tests
{
    public class MapFloatTests
    {
        [Fact]
        public void ReadRow_AutoMappedFloat_Success()
        {
            using (var importer = Helpers.GetImporter("Doubles.xlsx"))
            {
                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                // Valid cell value.
                FloatValue row1 = sheet.ReadRow<FloatValue>();
                Assert.Equal(2.2345f, row1.Value);

                // Empty cell value.
                Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<FloatValue>());

                // Invalid cell value.
                Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<FloatValue>());
            }
        }

        [Fact]
        public void ReadRow_AutoMappedNullableFloat_Success()
        {
            using (var importer = Helpers.GetImporter("Doubles.xlsx"))
            {
                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                // Valid cell value.
                NullableFloatValue row1 = sheet.ReadRow<NullableFloatValue>();
                Assert.Equal(2.2345f, row1.Value);

                // Empty cell value.
                NullableFloatValue row2 = sheet.ReadRow<NullableFloatValue>();
                Assert.Null(row2.Value);

                // Invalid cell value.
                Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableFloatValue>());
            }
        }

        [Fact]
        public void ReadRow_CustomMappedFloat_Success()
        {
            using (var importer = Helpers.GetImporter("Doubles.xlsx"))
            {
                importer.Configuration.RegisterClassMap<FloatValueFallbackMap>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                // Valid cell value.
                FloatValue row1 = sheet.ReadRow<FloatValue>();
                Assert.Equal(2.2345f, row1.Value);

                // Empty cell value.
                FloatValue row2 = sheet.ReadRow<FloatValue>();
                Assert.Equal(-10, row2.Value);

                // Invalid cell value.
                FloatValue row3 = sheet.ReadRow<FloatValue>();
                Assert.Equal(10, row3.Value);
            }
        }

        [Fact]
        public void ReadRow_CustomMappedNullableFloat_Success()
        {
            using (var importer = Helpers.GetImporter("Doubles.xlsx"))
            {
                importer.Configuration.RegisterClassMap<NullableFloatValueFallbackMap>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                // Valid cell value.
                NullableFloatValue row1 = sheet.ReadRow<NullableFloatValue>();
                Assert.Equal(2.2345f, row1.Value);

                // Empty cell value.
                NullableFloatValue row2 = sheet.ReadRow<NullableFloatValue>();
                Assert.Equal(-10, row2.Value);

                // Invalid cell value.
                NullableFloatValue row3 = sheet.ReadRow<NullableFloatValue>();
                Assert.Equal(10, row3.Value);
            }
        }

        private class FloatValue
        {
            public float Value { get; set; }
        }

        private class NullableFloatValue
        {
            public float? Value { get; set; }
        }

        private class FloatValueFallbackMap : ExcelClassMap<FloatValue>
        {
            public FloatValueFallbackMap()
            {
                Map(o => o.Value)
                    .WithEmptyFallback(-10.0f)
                    .WithInvalidFallback(10.0f);
            }
        }

        private class NullableFloatValueFallbackMap : ExcelClassMap<NullableFloatValue>
        {
            public NullableFloatValueFallbackMap()
            {
                Map(o => o.Value)
                    .WithEmptyFallback(-10.0f)
                    .WithInvalidFallback(10.0f);
            }
        }
    }
}

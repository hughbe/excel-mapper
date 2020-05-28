using Xunit;

namespace ExcelMapper.Tests
{
    public class MapFloatTests
    {
        [Fact]
        public void ReadRow_AutoMappedFloat_Success()
        {
            using var importer = Helpers.GetImporter("Doubles.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            FloatClass row1 = sheet.ReadRow<FloatClass>();
            Assert.Equal(2.2345f, row1.Value);

            // Empty cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<FloatClass>());

            // Invalid cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<FloatClass>());
        }

        [Fact]
        public void ReadRow_AutoMappedNullableFloat_Success()
        {
            using var importer = Helpers.GetImporter("Doubles.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            NullableFloatClass row1 = sheet.ReadRow<NullableFloatClass>();
            Assert.Equal(2.2345f, row1.Value);

            // Empty cell value.
            NullableFloatClass row2 = sheet.ReadRow<NullableFloatClass>();
            Assert.Null(row2.Value);

            // Invalid cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableFloatClass>());
        }
        [Fact]
        public void ReadRow_DefaultMappedFloat_Success()
        {
            using var importer = Helpers.GetImporter("Doubles.xlsx");
            importer.Configuration.RegisterClassMap<DefaultFloatClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            FloatClass row1 = sheet.ReadRow<FloatClass>();
            Assert.Equal(2.2345f, row1.Value);

            // Empty cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<FloatClass>());

            // Invalid cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<FloatClass>());
        }

        [Fact]
        public void ReadRow_DefaultMappedNullableFloat_Success()
        {
            using var importer = Helpers.GetImporter("Doubles.xlsx");
            importer.Configuration.RegisterClassMap<DefaultNullableFloatClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            NullableFloatClass row1 = sheet.ReadRow<NullableFloatClass>();
            Assert.Equal(2.2345f, row1.Value);

            // Empty cell value.
            NullableFloatClass row2 = sheet.ReadRow<NullableFloatClass>();
            Assert.Null(row2.Value);

            // Invalid cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableFloatClass>());
        }

        [Fact]
        public void ReadRow_CustomMappedFloat_Success()
        {
            using var importer = Helpers.GetImporter("Doubles.xlsx");
            importer.Configuration.RegisterClassMap<CustomFloatClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            FloatClass row1 = sheet.ReadRow<FloatClass>();
            Assert.Equal(2.2345f, row1.Value);

            // Empty cell value.
            FloatClass row2 = sheet.ReadRow<FloatClass>();
            Assert.Equal(-10, row2.Value);

            // Invalid cell value.
            FloatClass row3 = sheet.ReadRow<FloatClass>();
            Assert.Equal(10, row3.Value);
        }

        [Fact]
        public void ReadRow_CustomMappedNullableFloat_Success()
        {
            using var importer = Helpers.GetImporter("Doubles.xlsx");
            importer.Configuration.RegisterClassMap<CustomNullableFlatClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            NullableFloatClass row1 = sheet.ReadRow<NullableFloatClass>();
            Assert.Equal(2.2345f, row1.Value);

            // Empty cell value.
            NullableFloatClass row2 = sheet.ReadRow<NullableFloatClass>();
            Assert.Equal(-10, row2.Value);

            // Invalid cell value.
            NullableFloatClass row3 = sheet.ReadRow<NullableFloatClass>();
            Assert.Equal(10, row3.Value);
        }

        private class FloatClass
        {
            public float Value { get; set; }
        }

        private class DefaultFloatClassMap : ExcelClassMap<FloatClass>
        {
            public DefaultFloatClassMap()
            {
                Map(o => o.Value);
            }
        }

        private class CustomFloatClassMap : ExcelClassMap<FloatClass>
        {
            public CustomFloatClassMap()
            {
                Map(o => o.Value)
                    .WithEmptyFallback(-10.0f)
                    .WithInvalidFallback(10.0f);
            }
        }

        private class NullableFloatClass
        {
            public float? Value { get; set; }
        }

        private class DefaultNullableFloatClassMap : ExcelClassMap<NullableFloatClass>
        {
            public DefaultNullableFloatClassMap()
            {
                Map(o => o.Value);
            }
        }

        private class CustomNullableFlatClassMap : ExcelClassMap<NullableFloatClass>
        {
            public CustomNullableFlatClassMap()
            {
                Map(o => o.Value)
                    .WithEmptyFallback(-10.0f)
                    .WithInvalidFallback(10.0f);
            }
        }
    }
}

using Xunit;

namespace ExcelMapper.Tests
{
    public class MapDoubleTests
    {
        [Fact]
        public void ReadRow_AutoMappedDouble_Success()
        {
            using var importer = Helpers.GetImporter("Doubles.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            DoubleClass row1 = sheet.ReadRow<DoubleClass>();
            Assert.Equal(2.2345, row1.Value);

            // Empty cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DoubleClass>());

            // Invalid cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DoubleClass>());
        }

        [Fact]
        public void ReadRow_AutoMappedNullableDouble_Success()
        {
            using var importer = Helpers.GetImporter("Doubles.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            NullableDoubleClass row1 = sheet.ReadRow<NullableDoubleClass>();
            Assert.Equal(2.2345, row1.Value);

            // Empty cell value.
            NullableDoubleClass row2 = sheet.ReadRow<NullableDoubleClass>();
            Assert.Null(row2.Value);

            // Invalid cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableDoubleClass>());
        }

        [Fact]
        public void ReadRow_DefaultMappedDouble_Success()
        {
            using var importer = Helpers.GetImporter("Doubles.xlsx");
            importer.Configuration.RegisterClassMap<DefaultDoubleClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            DoubleClass row1 = sheet.ReadRow<DoubleClass>();
            Assert.Equal(2.2345, row1.Value);

            // Empty cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DoubleClass>());

            // Invalid cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DoubleClass>());
        }

        [Fact]
        public void ReadRow_DefaultMappedNullableDouble_Success()
        {
            using var importer = Helpers.GetImporter("Doubles.xlsx");
            importer.Configuration.RegisterClassMap<DefaultNullableDoubleClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            NullableDoubleClass row1 = sheet.ReadRow<NullableDoubleClass>();
            Assert.Equal(2.2345, row1.Value);

            // Empty cell value.
            NullableDoubleClass row2 = sheet.ReadRow<NullableDoubleClass>();
            Assert.Null(row2.Value);

            // Invalid cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableDoubleClass>());
        }

        [Fact]
        public void ReadRow_CustomMappedDouble_Success()
        {
            using var importer = Helpers.GetImporter("Doubles.xlsx");
            importer.Configuration.RegisterClassMap<CustomDoubleClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            DoubleClass row1 = sheet.ReadRow<DoubleClass>();
            Assert.Equal(2.2345, row1.Value);

            // Empty cell value.
            DoubleClass row2 = sheet.ReadRow<DoubleClass>();
            Assert.Equal(-10, row2.Value);

            // Invalid cell value.
            DoubleClass row3 = sheet.ReadRow<DoubleClass>();
            Assert.Equal(10, row3.Value);
        }

        [Fact]
        public void ReadRow_CustomMappedNullableDouble_Success()
        {
            using var importer = Helpers.GetImporter("Doubles.xlsx");
            importer.Configuration.RegisterClassMap<CustomNullableDoubleClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            NullableDoubleClass row1 = sheet.ReadRow<NullableDoubleClass>();
            Assert.Equal(2.2345, row1.Value);

            // Empty cell value.
            NullableDoubleClass row2 = sheet.ReadRow<NullableDoubleClass>();
            Assert.Equal(-10, row2.Value);

            // Invalid cell value.
            NullableDoubleClass row3 = sheet.ReadRow<NullableDoubleClass>();
            Assert.Equal(10, row3.Value);
        }

        private class DoubleClass
        {
            public double Value { get; set; }
        }

        private class DefaultDoubleClassMap : ExcelClassMap<DoubleClass>
        {
            public DefaultDoubleClassMap()
            {
                Map(o => o.Value);
            }
        }

        private class CustomDoubleClassMap : ExcelClassMap<DoubleClass>
        {
            public CustomDoubleClassMap()
            {
                Map(o => o.Value)
                    .WithEmptyFallback(-10.0)
                    .WithInvalidFallback(10.0);
            }
        }

        private class NullableDoubleClass
        {
            public double? Value { get; set; }
        }

        private class DefaultNullableDoubleClassMap : ExcelClassMap<NullableDoubleClass>
        {
            public DefaultNullableDoubleClassMap()
            {
                Map(o => o.Value);
            }
        }

        private class CustomNullableDoubleClassMap : ExcelClassMap<NullableDoubleClass>
        {
            public CustomNullableDoubleClassMap()
            {
                Map(o => o.Value)
                    .WithEmptyFallback(-10.0)
                    .WithInvalidFallback(10.0);
            }
        }
    }
}

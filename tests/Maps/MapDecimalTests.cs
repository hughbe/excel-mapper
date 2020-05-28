using Xunit;

namespace ExcelMapper.Tests
{
    public class MapDecimalTests
    {
        [Fact]
        public void ReadRow_AutoMappedDecimal_Success()
        {
            using var importer = Helpers.GetImporter("Doubles.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            DecimalClass row1 = sheet.ReadRow<DecimalClass>();
            Assert.Equal(2.2345m, row1.Value);

            // Empty cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DecimalClass>());

            // Invalid cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DecimalClass>());
        }

        [Fact]
        public void ReadRow_AutoMappedNullableDecimal_Success()
        {
            using var importer = Helpers.GetImporter("Doubles.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            NullableDecimalClass row1 = sheet.ReadRow<NullableDecimalClass>();
            Assert.Equal(2.2345m, row1.Value);

            // Empty cell value.
            NullableDecimalClass row2 = sheet.ReadRow<NullableDecimalClass>();
            Assert.Null(row2.Value);

            // Invalid cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DecimalClass>());
        }
        [Fact]
        public void ReadRow_DefaultMappedDecimal_Success()
        {
            using var importer = Helpers.GetImporter("Doubles.xlsx");
            importer.Configuration.RegisterClassMap<DefaultDecimalClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            DecimalClass row1 = sheet.ReadRow<DecimalClass>();
            Assert.Equal(2.2345m, row1.Value);

            // Empty cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DecimalClass>());

            // Invalid cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DecimalClass>());
        }

        [Fact]
        public void ReadRow_DefaultMappedNullableDecimal_Success()
        {
            using var importer = Helpers.GetImporter("Doubles.xlsx");
            importer.Configuration.RegisterClassMap<DefaultNullableDecimalClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            NullableDecimalClass row1 = sheet.ReadRow<NullableDecimalClass>();
            Assert.Equal(2.2345m, row1.Value);

            // Empty cell value.
            NullableDecimalClass row2 = sheet.ReadRow<NullableDecimalClass>();
            Assert.Null(row2.Value);

            // Invalid cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<DecimalClass>());
        }

        [Fact]
        public void ReadRow_CustomMappedDecimal_Success()
        {
            using var importer = Helpers.GetImporter("Doubles.xlsx");
            importer.Configuration.RegisterClassMap<CustomDecimalClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            DecimalClass row1 = sheet.ReadRow<DecimalClass>();
            Assert.Equal(2.2345m, row1.Value);

            // Empty cell value.
            DecimalClass row2 = sheet.ReadRow<DecimalClass>();
            Assert.Equal(-10, row2.Value);

            // Invalid cell value.
            DecimalClass row3 = sheet.ReadRow<DecimalClass>();
            Assert.Equal(10, row3.Value);
        }

        [Fact]
        public void ReadRow_CustomMappedNullableDecimal_Success()
        {
            using var importer = Helpers.GetImporter("Doubles.xlsx");
            importer.Configuration.RegisterClassMap<CustomNullableDecimalClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            NullableDecimalClass row1 = sheet.ReadRow<NullableDecimalClass>();
            Assert.Equal(2.2345m, row1.Value);

            // Empty cell value.
            NullableDecimalClass row2 = sheet.ReadRow<NullableDecimalClass>();
            Assert.Equal(-10, row2.Value);

            // Invalid cell value.
            NullableDecimalClass row3 = sheet.ReadRow<NullableDecimalClass>();
            Assert.Equal(10, row3.Value);
        }

        [Fact]
        public void ReadRow_CustomNullMappedNullableDecimal_Success()
        {
            using var importer = Helpers.GetImporter("Doubles.xlsx");
            importer.Configuration.RegisterClassMap<NullableDecimalNullValueFallbackMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            NullableDecimalClass row1 = sheet.ReadRow<NullableDecimalClass>();
            Assert.Equal(2.2345m, row1.Value);

            // Empty cell value.
            NullableDecimalClass row2 = sheet.ReadRow<NullableDecimalClass>();
            Assert.Equal((decimal?)null, row2.Value);

            // Invalid cell value.
            NullableDecimalClass row3 = sheet.ReadRow<NullableDecimalClass>();
            Assert.Equal((decimal?)null, row3.Value);
        }

        private class DecimalClass
        {
            public decimal Value { get; set; }
        }

        private class DefaultDecimalClassMap : ExcelClassMap<DecimalClass>
        {
            public DefaultDecimalClassMap()
            {
                Map(o => o.Value);
            }
        }

        private class CustomDecimalClassMap : ExcelClassMap<DecimalClass>
        {
            public CustomDecimalClassMap()
            {
                Map(o => o.Value)
                    .WithEmptyFallback(-10m)
                    .WithInvalidFallback(10m);
            }
        }

        private class NullableDecimalClass
        {
            public decimal? Value { get; set; }
        }

        private class DefaultNullableDecimalClassMap : ExcelClassMap<NullableDecimalClass>
        {
            public DefaultNullableDecimalClassMap()
            {
                Map(o => o.Value);
            }
        }

        private class CustomNullableDecimalClassMap : ExcelClassMap<NullableDecimalClass>
        {
            public CustomNullableDecimalClassMap()
            {
                Map(o => o.Value)
                    .WithEmptyFallback(-10m)
                    .WithInvalidFallback(10m);
            }
        }

        private class NullableDecimalNullValueFallbackMap : ExcelClassMap<NullableDecimalClass>
        {
            public NullableDecimalNullValueFallbackMap()
            {
                Map(o => o.Value)
                    .WithEmptyFallback((decimal?)null)
                    .WithInvalidFallback((decimal?)null);
            }
        }
    }
}

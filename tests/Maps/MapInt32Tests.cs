using Xunit;

namespace ExcelMapper.Tests
{
    public class MapInt32Tests
    {
        [Fact]
        public void ReadRow_AutoMappedInt32_Success()
        {
            using var importer = Helpers.GetImporter("Numbers.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            Int32Value row1 = sheet.ReadRow<Int32Value>();
            Assert.Equal(2, row1.Value);

            // Empty cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Int32Value>());

            // Invalid cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Int32Value>());
        }

        [Fact]
        public void ReadRow_AutoMappedNullableInt32_Success()
        {
            using var importer = Helpers.GetImporter("Numbers.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            NullableInt32Class row1 = sheet.ReadRow<NullableInt32Class>();
            Assert.Equal(2, row1.Value);

            // Empty cell value.
            NullableInt32Class row2 = sheet.ReadRow<NullableInt32Class>();
            Assert.Null(row2.Value);

            // Invalid cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Int32Value>());
        }

        [Fact]
        public void ReadRow_DefaultMappedInt32_Success()
        {
            using var importer = Helpers.GetImporter("Numbers.xlsx");
            importer.Configuration.RegisterClassMap<DefaultInt32ValueMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            Int32Value row1 = sheet.ReadRow<Int32Value>();
            Assert.Equal(2, row1.Value);

            // Empty cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Int32Value>());

            // Invalid cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Int32Value>());
        }

        [Fact]
        public void ReadRow_DefaultMappedNullableInt32_Success()
        {
            using var importer = Helpers.GetImporter("Numbers.xlsx");
            importer.Configuration.RegisterClassMap<DefaultNullableInt32ClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            NullableInt32Class row1 = sheet.ReadRow<NullableInt32Class>();
            Assert.Equal(2, row1.Value);

            // Empty cell value.
            NullableInt32Class row2 = sheet.ReadRow<NullableInt32Class>();
            Assert.Null(row2.Value);

            // Invalid cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<Int32Value>());
        }

        [Fact]
        public void ReadRow_CustomMappedInt32_Success()
        {
            using var importer = Helpers.GetImporter("Numbers.xlsx");
            importer.Configuration.RegisterClassMap<CustomInt32ValueMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            Int32Value row1 = sheet.ReadRow<Int32Value>();
            Assert.Equal(2, row1.Value);

            // Empty cell value.
            Int32Value row2 = sheet.ReadRow<Int32Value>();
            Assert.Equal(-10, row2.Value);

            // Invalid cell value.
            Int32Value row3 = sheet.ReadRow<Int32Value>();
            Assert.Equal(10, row3.Value);
        }

        [Fact]
        public void ReadRow_CustomMappedNullableInt32_Success()
        {
            using var importer = Helpers.GetImporter("Numbers.xlsx");
            importer.Configuration.RegisterClassMap<CustomNullableInt32ClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            NullableInt32Class row1 = sheet.ReadRow<NullableInt32Class>();
            Assert.Equal(2, row1.Value);

            // Empty cell value.
            NullableInt32Class row2 = sheet.ReadRow<NullableInt32Class>();
            Assert.Equal(-10, row2.Value);

            // Invalid cell value.
            NullableInt32Class row3 = sheet.ReadRow<NullableInt32Class>();
            Assert.Equal(10, row3.Value);
        }

        private class Int32Value
        {
            public int Value { get; set; }
        }

        private class DefaultInt32ValueMap : ExcelClassMap<Int32Value>
        {
            public DefaultInt32ValueMap()
            {
                Map(o => o.Value);
            }
        }

        private class CustomInt32ValueMap : ExcelClassMap<Int32Value>
        {
            public CustomInt32ValueMap()
            {
                Map(o => o.Value)
                    .WithEmptyFallback(-10)
                    .WithInvalidFallback(10);
            }
        }

        private class NullableInt32Class
        {
            public int? Value { get; set; }
        }

        private class DefaultNullableInt32ClassMap : ExcelClassMap<NullableInt32Class>
        {
            public DefaultNullableInt32ClassMap()
            {
                Map(o => o.Value);
            }
        }

        private class CustomNullableInt32ClassMap : ExcelClassMap<NullableInt32Class>
        {
            public CustomNullableInt32ClassMap()
            {
                Map(o => o.Value)
                    .WithEmptyFallback(-10)
                    .WithInvalidFallback(10);
            }
        }
    }
}

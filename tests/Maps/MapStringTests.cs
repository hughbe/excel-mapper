using Xunit;

namespace ExcelMapper.Tests
{
    public class MapStringTests
    {
        [Fact]
        public void ReadRow_AutoMappedString_Success()
        {
            using var importer = Helpers.GetImporter("Strings.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid value
            StringClass row1 = sheet.ReadRow<StringClass>();
            Assert.Equal("value", row1.Value);

            // Valid value
            StringClass row2 = sheet.ReadRow<StringClass>();
            Assert.Equal("  value  ", row2.Value);

            // Empty value
            StringClass row3 = sheet.ReadRow<StringClass>();
            Assert.Null(row3.Value);
        }

        [Fact]
        public void ReadRow_DefaultMappedString_Success()
        {
            using var importer = Helpers.GetImporter("Strings.xlsx");
            importer.Configuration.RegisterClassMap<DefaultStringClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid value
            StringClass row1 = sheet.ReadRow<StringClass>();
            Assert.Equal("value", row1.Value);

            // Valid value
            StringClass row2 = sheet.ReadRow<StringClass>();
            Assert.Equal("  value  ", row2.Value);

            // Empty value
            StringClass row3 = sheet.ReadRow<StringClass>();
            Assert.Null(row3.Value);
        }

        [Fact]
        public void ReadRow_CustomMappedString_Success()
        {
            using var importer = Helpers.GetImporter("Strings.xlsx");
            importer.Configuration.RegisterClassMap<CustomStringClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid value
            StringClass row1 = sheet.ReadRow<StringClass>();
            Assert.Equal("value", row1.Value);

            // Valid value
            StringClass row2 = sheet.ReadRow<StringClass>();
            Assert.Equal("  value  ", row2.Value);

            // Empty value
            StringClass row3 = sheet.ReadRow<StringClass>();
            Assert.Equal("empty", row3.Value);
        }

        private class StringClass
        {
            public string Value { get; set; }
        }

        private class DefaultStringClassMap : ExcelClassMap<StringClass>
        {
            public DefaultStringClassMap()
            {
                Map(o => o.Value);
            }
        }

        private class CustomStringClassMap : ExcelClassMap<StringClass>
        {
            public CustomStringClassMap()
            {
                Map(o => o.Value)
                    .WithEmptyFallback("empty")
                    .WithInvalidFallback("invalid");
            }
        }
    }
}

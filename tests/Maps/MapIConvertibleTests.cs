using System;
using Xunit;

namespace ExcelMapper.Tests
{
    public class MapIConvertibleTests
    {
        [Fact]
        public void ReadRow_AutoMappedIConvertible_Success()
        {
            using var importer = Helpers.GetImporter("Strings.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid value
            ConvertibleClass row1 = sheet.ReadRow<ConvertibleClass>();
            Assert.Equal("value", row1.Value);

            // Valid value
            ConvertibleClass row2 = sheet.ReadRow<ConvertibleClass>();
            Assert.Equal("  value  ", row2.Value);

            // Empty value
            ConvertibleClass row3 = sheet.ReadRow<ConvertibleClass>();
            Assert.Null(row3.Value);
        }

        [Fact]
        public void ReadRow_DefaultMappedIConvertible_Success()
        {
            using var importer = Helpers.GetImporter("Strings.xlsx");
            importer.Configuration.RegisterClassMap<DefaultConvertibleClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid value
            ConvertibleClass row1 = sheet.ReadRow<ConvertibleClass>();
            Assert.Equal("value", row1.Value);

            // Valid value
            ConvertibleClass row2 = sheet.ReadRow<ConvertibleClass>();
            Assert.Equal("  value  ", row2.Value);

            // Empty value
            ConvertibleClass row3 = sheet.ReadRow<ConvertibleClass>();
            Assert.Null(row3.Value);
        }

        [Fact]
        public void ReadRow_CustomMappedIConvertible_Success()
        {
            using var importer = Helpers.GetImporter("Strings.xlsx");
            importer.Configuration.RegisterClassMap<CustomConvertibleClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid value
            ConvertibleClass row1 = sheet.ReadRow<ConvertibleClass>();
            Assert.Equal("value", row1.Value);

            // Valid value
            ConvertibleClass row2 = sheet.ReadRow<ConvertibleClass>();
            Assert.Equal("  value  ", row2.Value);

            // Empty value
            ConvertibleClass row3 = sheet.ReadRow<ConvertibleClass>();
            Assert.Equal("empty", row3.Value);
        }

        private class ConvertibleClass
        {
            public IConvertible Value { get; set; }
        }

        private class DefaultConvertibleClassMap : ExcelClassMap<ConvertibleClass>
        {
            public DefaultConvertibleClassMap()
            {
                Map(o => o.Value);
            }
        }

        private class CustomConvertibleClassMap : ExcelClassMap<ConvertibleClass>
        {
            public CustomConvertibleClassMap()
            {
                Map(o => o.Value)
                    .WithEmptyFallback("empty")
                    .WithInvalidFallback("invalid");
            }
        }
    }
}

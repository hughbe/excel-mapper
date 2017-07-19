using System;
using Xunit;

namespace ExcelMapper.Tests
{
    public class MapIConvertibleTests
    {
        [Fact]
        public void ReadRow_AutoMappedIConvertible_Success()
        {
            using (var importer = Helpers.GetImporter("Strings.xlsx"))
            {
                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                // Valid value
                ConvertibleValue row1 = sheet.ReadRow<ConvertibleValue>();
                Assert.Equal("value", row1.Value);

                // Valid value
                ConvertibleValue row2 = sheet.ReadRow<ConvertibleValue>();
                Assert.Equal("  value  ", row2.Value);

                // Empty value
                ConvertibleValue row3 = sheet.ReadRow<ConvertibleValue>();
                Assert.Null(row3.Value);
            }
        }

        [Fact]
        public void ReadRow_CustomMappedIConvertible_Success()
        {
            using (var importer = Helpers.GetImporter("Strings.xlsx"))
            {
                importer.Configuration.RegisterClassMap<ConvertibleValueFallbackMap>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                // Valid value
                ConvertibleValue row1 = sheet.ReadRow<ConvertibleValue>();
                Assert.Equal("value", row1.Value);

                // Valid value
                ConvertibleValue row2 = sheet.ReadRow<ConvertibleValue>();
                Assert.Equal("  value  ", row2.Value);

                // Empty value
                ConvertibleValue row3 = sheet.ReadRow<ConvertibleValue>();
                Assert.Equal("empty", row3.Value);
            }
        }

        private class ConvertibleValue
        {
            public IConvertible Value { get; set; }
        }

        private class ConvertibleValueFallbackMap : ExcelClassMap<ConvertibleValue>
        {
            public ConvertibleValueFallbackMap()
            {
                Map(o => o.Value)
                    .WithEmptyFallback("empty")
                    .WithInvalidFallback("invalid");
            }
        }
    }
}

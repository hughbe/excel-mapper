using System;
using Xunit;

namespace ExcelMapper.Tests
{
    public class MapStringTests
    {
        [Fact]
        public void ReadRow_AutoMappedString_Success()
        {
            using (var importer = Helpers.GetImporter("Strings.xlsx"))
            {
                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                // Valid value
                StringValue row1 = sheet.ReadRow<StringValue>();
                Assert.Equal("value", row1.Value);

                // Valid value
                StringValue row2 = sheet.ReadRow<StringValue>();
                Assert.Equal("  value  ", row2.Value);

                // Empty value
                StringValue row3 = sheet.ReadRow<StringValue>();
                Assert.Null(row3.Value);
            }
        }

        [Fact]
        public void ReadRow_CustomMappedString_Success()
        {
            using (var importer = Helpers.GetImporter("Strings.xlsx"))
            {
                importer.Configuration.RegisterClassMap<StringValueFallbackMap>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                // Valid value
                StringValue row1 = sheet.ReadRow<StringValue>();
                Assert.Equal("value", row1.Value);

                // Valid value
                StringValue row2 = sheet.ReadRow<StringValue>();
                Assert.Equal("  value  ", row2.Value);

                // Empty value
                StringValue row3 = sheet.ReadRow<StringValue>();
                Assert.Equal("empty", row3.Value);
            }
        }

        private class StringValue
        {
            public string Value { get; set; }
        }

        private class StringValueFallbackMap : ExcelClassMap<StringValue>
        {
            public StringValueFallbackMap()
            {
                Map(o => o.Value)
                    .WithEmptyFallback("empty")
                    .WithInvalidFallback("invalid");
            }
        }
    }
}

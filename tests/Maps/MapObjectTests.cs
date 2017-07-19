using Xunit;

namespace ExcelMapper.Tests
{
    public class MapObjectTests
    {
        [Fact]
        public void ReadRow_AutoMappedObject_Success()
        {
            using (var importer = Helpers.GetImporter("Strings.xlsx"))
            {
                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                // Valid value
                ObjectValue row1 = sheet.ReadRow<ObjectValue>();
                Assert.Equal("value", row1.Value);

                // Valid value
                ObjectValue row2 = sheet.ReadRow<ObjectValue>();
                Assert.Equal("  value  ", row2.Value);

                // Empty value
                ObjectValue row3 = sheet.ReadRow<ObjectValue>();
                Assert.Null(row3.Value);
            }
        }

        [Fact]
        public void ReadRow_CustomMappedObject_Success()
        {
            using (var importer = Helpers.GetImporter("Strings.xlsx"))
            {
                importer.Configuration.RegisterClassMap<ObjectValueFallbackMap>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                // Valid value
                ObjectValue row1 = sheet.ReadRow<ObjectValue>();
                Assert.Equal("value", row1.Value);

                // Valid value
                ObjectValue row2 = sheet.ReadRow<ObjectValue>();
                Assert.Equal("  value  ", row2.Value);

                // Empty value
                ObjectValue row3 = sheet.ReadRow<ObjectValue>();
                Assert.Equal("empty", row3.Value);
            }
        }

        private class ObjectValue
        {
            public object Value { get; set; }
        }

        private class ObjectValueFallbackMap : ExcelClassMap<ObjectValue>
        {
            public ObjectValueFallbackMap()
            {
                Map(o => o.Value)
                    .WithEmptyFallback("empty")
                    .WithInvalidFallback("invalid");
            }
        }
    }
}

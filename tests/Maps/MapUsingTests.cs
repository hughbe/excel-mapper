using Xunit;

namespace ExcelMapper.Tests
{
    public class MapUsingTests
    {
        [Fact]
        public void ReadRow_ConvertUsingMap_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                importer.Configuration.RegisterClassMap<ConvertUsingValueMap>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                ConvertUsingValue row1 = sheet.ReadRow<ConvertUsingValue>();
                Assert.Equal("aextra", row1.StringValue);
            }
        }

        private class ConvertUsingValue
        {
            public string StringValue { get; set; }
        }

        private class ConvertUsingValueMap : ExcelClassMap<ConvertUsingValue>
        {
            public ConvertUsingValueMap()
            {
                Map(c => c.StringValue)
                    .WithConverter(s => s + "extra");
            }
        }
    }
}

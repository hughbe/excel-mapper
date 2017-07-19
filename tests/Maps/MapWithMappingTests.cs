using System;
using System.Collections.Generic;
using Xunit;

namespace ExcelMapper.Tests
{
    public class MapWithMappingTests
    {
        [Fact]
        public void ReadRow_WithMappingMap_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("WithMappings.xlsx"))
            {
                importer.Configuration.RegisterClassMap<WithMappingValueMap>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                WithMappingValue row1 = sheet.ReadRow<WithMappingValue>();
                Assert.Equal("a", row1.StringValue);
                Assert.Equal(MapUsingValueEnum.First, row1.EnumValue);

                WithMappingValue row2 = sheet.ReadRow<WithMappingValue>();
                Assert.Equal("extra", row2.StringValue);
                Assert.Equal(MapUsingValueEnum.First, row2.EnumValue);

                WithMappingValue row3 = sheet.ReadRow<WithMappingValue>();
                Assert.Equal("extra", row3.StringValue);
                Assert.Equal(MapUsingValueEnum.Second, row3.EnumValue);

                WithMappingValue row4 = sheet.ReadRow<WithMappingValue>();
                Assert.Null(row4.StringValue);
                Assert.Equal(MapUsingValueEnum.Unknown, row4.EnumValue);
            }
        }

        private enum MapUsingValueEnum
        {
            First,
            Second,
            Unknown
        }

        private class WithMappingValue
        {
            public string StringValue { get; set; }
            public MapUsingValueEnum EnumValue { get; set; }
        }

        private class WithMappingValueMap : ExcelClassMap<WithMappingValue>
        {
            public WithMappingValueMap()
            {
                Map(c => c.StringValue)
                    .WithMapping(new Dictionary<string, string>
                    {
                        { "b", "extra" }
                    }, StringComparer.OrdinalIgnoreCase);

                Map(c => c.EnumValue)
                    .WithMapping(new Dictionary<string, MapUsingValueEnum>
                    {
                        { "one", MapUsingValueEnum.First }
                    })
                    .WithInvalidFallback(MapUsingValueEnum.Unknown);
            }
        }
    }
}

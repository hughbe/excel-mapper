using Xunit;

namespace ExcelMapper.Tests
{
    public class MatchingMapTests
    {
        [Fact]
        public void ReadRow_HasHeader_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                importer.Configuration.RegisterClassMap<ColumnMatchingClassMap>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                StringValue row1 = sheet.ReadRow<StringValue>();
                Assert.Equal("1", row1.Value);

                StringValue row2 = sheet.ReadRow<StringValue>();
                Assert.Equal("b", row2.Value);

                StringValue row3 = sheet.ReadRow<StringValue>();
                Assert.Null(row3.Value);
            }
        }

        [Fact]
        public void ReadRow_NoHeader_ThrowsExcelMappingException()
        {
            using (var importer = Helpers.GetImporter("Non Zero Header Index.xlsx"))
            {
                importer.Configuration.RegisterClassMap<ColumnMatchingClassMap>();

                ExcelSheet sheet = importer.ReadSheet();

                Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<StringValue>());
            }
        }

        [Fact]
        public void ReadRow_NothingMatchingNotOptional_ThrowsExcelMappingException()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                importer.Configuration.RegisterClassMap<NoColumnMatchingClassMap>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<StringValue>());
            }
        }

        [Fact]
        public void ReadRow_NothingMatchingOptional_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                importer.Configuration.RegisterClassMap<NoColumnOptionalMatchingClassMap>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                StringValue row1 = sheet.ReadRow<StringValue>();
                Assert.Null(row1.Value);
            }
        }

        private class StringValue
        {
            public string Value { get; set; }
        }

        private class ColumnMatchingClassMap : ExcelClassMap<StringValue>
        {
            public ColumnMatchingClassMap()
            {
                int i = 0;

                Map(o => o.Value)
                    .WithColumnNameMatching(s =>
                    {
                        if (s == "Int Value" && i == 0)
                        {
                            i++;
                            return true;
                        }

                        if (s == "StringValue")
                        {
                            i++;
                            return true;
                        }

                        return false;
                    });
            }
        }

        private class NoColumnMatchingClassMap : ExcelClassMap<StringValue>
        {
            public NoColumnMatchingClassMap()
            {
                Map(o => o.Value)
                    .WithColumnNameMatching(s => false);
            }
        }

        private class NoColumnOptionalMatchingClassMap : ExcelClassMap<StringValue>
        {
            public NoColumnOptionalMatchingClassMap()
            {
                Map(o => o.Value)
                    .WithColumnNameMatching(s => false)
                    .MakeOptional();
            }
        }
    }
}

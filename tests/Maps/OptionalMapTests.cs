using Xunit;

namespace ExcelMapper.Tests
{
    public class OptionalMapTests
    {
        [Fact]
        public void ReadRow_OptionalMap_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                importer.Configuration.RegisterClassMap<OptionalValueMap>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                OptionalValue row1 = sheet.ReadRow<OptionalValue>();
                Assert.Equal(0, row1.NoSuchColumnNoName);
                Assert.Equal(0, row1.NoSuchColumnWithNameBefore);
                Assert.Equal(0, row1.NoSuchColumnWithNameAfter);
                Assert.Equal(0, row1.NoSuchColumnWithIndexBefore);
                Assert.Equal(0, row1.NoSuchColumnWithIndexAfter);
            }
        }

        private class OptionalValue
        {
            public int NoSuchColumnNoName { get; set; }

            public int NoSuchColumnWithNameBefore { get; set; }
            public int NoSuchColumnWithNameAfter { get; set; }

            public int NoSuchColumnWithIndexBefore { get; set; }
            public int NoSuchColumnWithIndexAfter { get; set; }
        }

        private class OptionalValueMap : ExcelClassMap<OptionalValue>
        {
            public OptionalValueMap()
            {
                Map(v => v.NoSuchColumnNoName)
                    .MakeOptional()
                    .WithEmptyFallback(-1);

                Map(v => v.NoSuchColumnWithNameBefore)
                    .WithColumnName("NoSuchColumn")
                    .MakeOptional()
                    .WithEmptyFallback(-2);

                Map(v => v.NoSuchColumnWithNameAfter)
                    .MakeOptional()
                    .WithColumnName("NoSuchColumn")
                    .WithEmptyFallback(-3);

                Map(v => v.NoSuchColumnWithIndexBefore)
                    .WithColumnIndex(10)
                    .MakeOptional()
                    .WithEmptyFallback(-4);

                Map(v => v.NoSuchColumnWithIndexAfter)
                    .MakeOptional()
                    .WithColumnIndex(10)
                    .WithEmptyFallback(-5);
            }
        }
    }
}

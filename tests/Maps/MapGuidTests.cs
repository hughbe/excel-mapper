using System;
using Xunit;

namespace ExcelMapper.Tests
{
    public class MapGuidTests
    {
        [Fact]
        public void ReadRow_AutoMappedGuid_Success()
        {
            using (var importer = Helpers.GetImporter("Guids.xlsx"))
            {
                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                // Valid cell value.
                GuidValue row1 = sheet.ReadRow<GuidValue>();
                Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f1"), row1.Value);

                GuidValue row2 = sheet.ReadRow<GuidValue>();
                Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f2"), row2.Value);

                GuidValue row3 = sheet.ReadRow<GuidValue>();
                Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f3"), row3.Value);

                GuidValue row4 = sheet.ReadRow<GuidValue>();
                Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f4"), row4.Value);

                GuidValue row5 = sheet.ReadRow<GuidValue>();
                Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f5"), row5.Value);

                // Empty cell value.
                Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<GuidValue>());

                // Invalid cell value.
                Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<GuidValue>());
            }
        }

        [Fact]
        public void ReadRow_AutoMappedNullableGuid_Success()
        {
            using (var importer = Helpers.GetImporter("Guids.xlsx"))
            {
                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                // Valid cell value.
                NullableGuidValue row1 = sheet.ReadRow<NullableGuidValue>();
                Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f1"), row1.Value);

                NullableGuidValue row2 = sheet.ReadRow<NullableGuidValue>();
                Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f2"), row2.Value);

                NullableGuidValue row3 = sheet.ReadRow<NullableGuidValue>();
                Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f3"), row3.Value);

                NullableGuidValue row4 = sheet.ReadRow<NullableGuidValue>();
                Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f4"), row4.Value);

                NullableGuidValue row5 = sheet.ReadRow<NullableGuidValue>();
                Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f5"), row5.Value);

                // Empty cell value.
                NullableGuidValue row6 = sheet.ReadRow<NullableGuidValue>();
                Assert.Null(row6.Value);

                // Invalid cell value.
                Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableGuidValue>());
            }
        }

        [Fact]
        public void ReadRow_CustomMappedGuid_Success()
        {
            using (var importer = Helpers.GetImporter("Guids.xlsx"))
            {
                importer.Configuration.RegisterClassMap<GuidValueFallbackMap>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                // Valid cell value.
                GuidValue row1 = sheet.ReadRow<GuidValue>();
                Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f1"), row1.Value);

                GuidValue row2 = sheet.ReadRow<GuidValue>();
                Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f2"), row2.Value);

                GuidValue row3 = sheet.ReadRow<GuidValue>();
                Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f3"), row3.Value);

                GuidValue row4 = sheet.ReadRow<GuidValue>();
                Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f4"), row4.Value);

                GuidValue row5 = sheet.ReadRow<GuidValue>();
                Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f5"), row5.Value);

                // Empty cell value.
                GuidValue row6 = sheet.ReadRow<GuidValue>();
                Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f6"), row6.Value);

                // Invalid cell value.
                GuidValue row7 = sheet.ReadRow<GuidValue>();
                Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f7"), row7.Value);
            }
        }

        [Fact]
        public void ReadRow_CustomMappedNullableGuid_Success()
        {
            using (var importer = Helpers.GetImporter("Guids.xlsx"))
            {
                importer.Configuration.RegisterClassMap<NullableGuidValueFallbackMap>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                // Valid cell value.
                NullableGuidValue row1 = sheet.ReadRow<NullableGuidValue>();
                Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f1"), row1.Value);

                NullableGuidValue row2 = sheet.ReadRow<NullableGuidValue>();
                Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f2"), row2.Value);

                NullableGuidValue row3 = sheet.ReadRow<NullableGuidValue>();
                Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f3"), row3.Value);

                NullableGuidValue row4 = sheet.ReadRow<NullableGuidValue>();
                Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f4"), row4.Value);

                NullableGuidValue row5 = sheet.ReadRow<NullableGuidValue>();
                Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f5"), row5.Value);

                // Empty cell value.
                NullableGuidValue row6 = sheet.ReadRow<NullableGuidValue>();
                Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f6"), row6.Value);

                // Invalid cell value.
                NullableGuidValue row7 = sheet.ReadRow<NullableGuidValue>();
                Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f7"), row7.Value);
            }
        }

        private class GuidValue
        {
            public Guid Value { get; set; }
        }

        private class NullableGuidValue
        {
            public Guid? Value { get; set; }
        }

        private class GuidValueFallbackMap : ExcelClassMap<GuidValue>
        {
            public GuidValueFallbackMap()
            {
                Map(o => o.Value)
                    .WithEmptyFallback(new Guid("a8a110d5fc4943c5bf46802db8f843f6"))
                    .WithInvalidFallback(new Guid("a8a110d5fc4943c5bf46802db8f843f7"));
            }
        }

        private class NullableGuidValueFallbackMap : ExcelClassMap<NullableGuidValue>
        {
            public NullableGuidValueFallbackMap()
            {
                Map(o => o.Value)
                    .WithEmptyFallback(new Guid("a8a110d5fc4943c5bf46802db8f843f6"))
                    .WithInvalidFallback(new Guid("a8a110d5fc4943c5bf46802db8f843f7"));
            }
        }
    }
}

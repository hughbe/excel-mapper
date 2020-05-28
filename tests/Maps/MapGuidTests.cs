using System;
using Xunit;

namespace ExcelMapper.Tests
{
    public class MapGuidTests
    {
        [Fact]
        public void ReadRow_AutoMappedGuid_Success()
        {
            using var importer = Helpers.GetImporter("Guids.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            GuidClass row1 = sheet.ReadRow<GuidClass>();
            Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f1"), row1.Value);

            GuidClass row2 = sheet.ReadRow<GuidClass>();
            Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f2"), row2.Value);

            GuidClass row3 = sheet.ReadRow<GuidClass>();
            Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f3"), row3.Value);

            GuidClass row4 = sheet.ReadRow<GuidClass>();
            Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f4"), row4.Value);

            GuidClass row5 = sheet.ReadRow<GuidClass>();
            Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f5"), row5.Value);

            // Empty cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<GuidClass>());

            // Invalid cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<GuidClass>());
        }

        [Fact]
        public void ReadRow_AutoMappedNullableGuid_Success()
        {
            using var importer = Helpers.GetImporter("Guids.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            NullableGuidClass row1 = sheet.ReadRow<NullableGuidClass>();
            Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f1"), row1.Value);

            NullableGuidClass row2 = sheet.ReadRow<NullableGuidClass>();
            Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f2"), row2.Value);

            NullableGuidClass row3 = sheet.ReadRow<NullableGuidClass>();
            Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f3"), row3.Value);

            NullableGuidClass row4 = sheet.ReadRow<NullableGuidClass>();
            Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f4"), row4.Value);

            NullableGuidClass row5 = sheet.ReadRow<NullableGuidClass>();
            Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f5"), row5.Value);

            // Empty cell value.
            NullableGuidClass row6 = sheet.ReadRow<NullableGuidClass>();
            Assert.Null(row6.Value);

            // Invalid cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableGuidClass>());
        }
        [Fact]
        public void ReadRow_DefaultMappedGuid_Success()
        {
            using var importer = Helpers.GetImporter("Guids.xlsx");
            importer.Configuration.RegisterClassMap<DefaultGuidClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            GuidClass row1 = sheet.ReadRow<GuidClass>();
            Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f1"), row1.Value);

            GuidClass row2 = sheet.ReadRow<GuidClass>();
            Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f2"), row2.Value);

            GuidClass row3 = sheet.ReadRow<GuidClass>();
            Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f3"), row3.Value);

            GuidClass row4 = sheet.ReadRow<GuidClass>();
            Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f4"), row4.Value);

            GuidClass row5 = sheet.ReadRow<GuidClass>();
            Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f5"), row5.Value);

            // Empty cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<GuidClass>());

            // Invalid cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<GuidClass>());
        }

        [Fact]
        public void ReadRow_DefaultMappedNullableGuid_Success()
        {
            using var importer = Helpers.GetImporter("Guids.xlsx");
            importer.Configuration.RegisterClassMap<DefaultNullableGuidClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            NullableGuidClass row1 = sheet.ReadRow<NullableGuidClass>();
            Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f1"), row1.Value);

            NullableGuidClass row2 = sheet.ReadRow<NullableGuidClass>();
            Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f2"), row2.Value);

            NullableGuidClass row3 = sheet.ReadRow<NullableGuidClass>();
            Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f3"), row3.Value);

            NullableGuidClass row4 = sheet.ReadRow<NullableGuidClass>();
            Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f4"), row4.Value);

            NullableGuidClass row5 = sheet.ReadRow<NullableGuidClass>();
            Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f5"), row5.Value);

            // Empty cell value.
            NullableGuidClass row6 = sheet.ReadRow<NullableGuidClass>();
            Assert.Null(row6.Value);

            // Invalid cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableGuidClass>());
        }

        [Fact]
        public void ReadRow_CustomMappedGuid_Success()
        {
            using var importer = Helpers.GetImporter("Guids.xlsx");
            importer.Configuration.RegisterClassMap<CustomGuidClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            GuidClass row1 = sheet.ReadRow<GuidClass>();
            Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f1"), row1.Value);

            GuidClass row2 = sheet.ReadRow<GuidClass>();
            Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f2"), row2.Value);

            GuidClass row3 = sheet.ReadRow<GuidClass>();
            Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f3"), row3.Value);

            GuidClass row4 = sheet.ReadRow<GuidClass>();
            Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f4"), row4.Value);

            GuidClass row5 = sheet.ReadRow<GuidClass>();
            Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f5"), row5.Value);

            // Empty cell value.
            GuidClass row6 = sheet.ReadRow<GuidClass>();
            Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f6"), row6.Value);

            // Invalid cell value.
            GuidClass row7 = sheet.ReadRow<GuidClass>();
            Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f7"), row7.Value);
        }

        [Fact]
        public void ReadRow_CustomMappedNullableGuid_Success()
        {
            using var importer = Helpers.GetImporter("Guids.xlsx");
            importer.Configuration.RegisterClassMap<CustomNullableGuidClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            NullableGuidClass row1 = sheet.ReadRow<NullableGuidClass>();
            Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f1"), row1.Value);

            NullableGuidClass row2 = sheet.ReadRow<NullableGuidClass>();
            Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f2"), row2.Value);

            NullableGuidClass row3 = sheet.ReadRow<NullableGuidClass>();
            Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f3"), row3.Value);

            NullableGuidClass row4 = sheet.ReadRow<NullableGuidClass>();
            Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f4"), row4.Value);

            NullableGuidClass row5 = sheet.ReadRow<NullableGuidClass>();
            Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f5"), row5.Value);

            // Empty cell value.
            NullableGuidClass row6 = sheet.ReadRow<NullableGuidClass>();
            Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f6"), row6.Value);

            // Invalid cell value.
            NullableGuidClass row7 = sheet.ReadRow<NullableGuidClass>();
            Assert.Equal(new Guid("a8a110d5fc4943c5bf46802db8f843f7"), row7.Value);
        }

        private class GuidClass
        {
            public Guid Value { get; set; }
        }

        private class DefaultGuidClassMap : ExcelClassMap<GuidClass>
        {
            public DefaultGuidClassMap()
            {
                Map(o => o.Value);
            }
        }

        private class CustomGuidClassMap : ExcelClassMap<GuidClass>
        {
            public CustomGuidClassMap()
            {
                Map(o => o.Value)
                    .WithEmptyFallback(new Guid("a8a110d5fc4943c5bf46802db8f843f6"))
                    .WithInvalidFallback(new Guid("a8a110d5fc4943c5bf46802db8f843f7"));
            }
        }

        private class NullableGuidClass
        {
            public Guid? Value { get; set; }
        }

        private class DefaultNullableGuidClassMap : ExcelClassMap<NullableGuidClass>
        {
            public DefaultNullableGuidClassMap()
            {
                Map(o => o.Value);
            }
        }

        private class CustomNullableGuidClassMap : ExcelClassMap<NullableGuidClass>
        {
            public CustomNullableGuidClassMap()
            {
                Map(o => o.Value)
                    .WithEmptyFallback(new Guid("a8a110d5fc4943c5bf46802db8f843f6"))
                    .WithInvalidFallback(new Guid("a8a110d5fc4943c5bf46802db8f843f7"));
            }
        }
    }
}

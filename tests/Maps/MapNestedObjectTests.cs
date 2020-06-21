using Xunit;

namespace ExcelMapper.Tests
{
    public class MapNestedObjectTests
    {
        [Fact]
        public void ReadRow_AutoMappedObject_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("NestedObjects.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            NestedObjectValue row1 = sheet.ReadRow<NestedObjectValue>();
            Assert.Equal("a", row1.SubValue1.StringValue);
            Assert.Equal(new string[] { "a", "b" }, row1.SubValue1.SplitStringValue);
            Assert.Equal(1, row1.SubValue2.IntValue);
            Assert.Equal(10, row1.SubValue2.SubValue.SubInt);
            Assert.Equal("c", row1.SubValue2.SubValue.SubString);
        }

        private class NestedObjectValue
        {
            public SubValue1 SubValue1 { get; set; }
            public SubValue2 SubValue2 { get; set; }
        }

        private class SubValue1
        {
            public string StringValue { get; set; }
            public string[] SplitStringValue { get; set; }
        }

        private class SubValue2
        {
            public int IntValue { get; set; }
            public SubValue3 SubValue { get; set; }
        }

        private class SubValue3
        {
            public string SubString { get; set; }
            public int SubInt { get; set; }
        }

        [Fact]
        public void ReadRow_CustomMappedObject_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("NestedObjects.xlsx");
            importer.Configuration.RegisterClassMap<ObjectValueCustomClassMapMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            NestedObjectValue row1 = sheet.ReadRow<NestedObjectValue>();
            Assert.Equal("a", row1.SubValue1.StringValue);
            Assert.Equal(new string[] { "a", "b" }, row1.SubValue1.SplitStringValue);
            Assert.Equal(1, row1.SubValue2.IntValue);
            Assert.Equal(10, row1.SubValue2.SubValue.SubInt);
            Assert.Equal("c", row1.SubValue2.SubValue.SubString);
        }

        private class ObjectValueCustomClassMapMap : ExcelClassMap<NestedObjectValue>
        {
            public ObjectValueCustomClassMapMap()
            {
                MapObject(p => p.SubValue1).WithClassMap(m =>
                {
                    m.Map(s => s.StringValue);
                    m.Map(s => s.SplitStringValue);
                });

                MapObject(p => p.SubValue2).WithClassMap(new SubValueMap());
            }
        }

        private class SubValueMap : ExcelClassMap<SubValue2>
        {
            public SubValueMap()
            {
                Map(s => s.IntValue);

                MapObject(s => s.SubValue);
            }
        }

        [Fact]
        public void ReadRow_CustomInnerObjectMap_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("NestedObjects.xlsx");
            importer.Configuration.RegisterClassMap<ObjectValueInnerMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            NestedObjectValue row1 = sheet.ReadRow<NestedObjectValue>();
            Assert.Equal("a", row1.SubValue1.StringValue);
            Assert.Equal(new string[] { "a", "b" }, row1.SubValue1.SplitStringValue);
            Assert.Equal(1, row1.SubValue2.IntValue);
            Assert.Equal(10, row1.SubValue2.SubValue.SubInt);
            Assert.Equal("c", row1.SubValue2.SubValue.SubString);
        }

        private class ObjectValueInnerMap : ExcelClassMap<NestedObjectValue>
        {
            public ObjectValueInnerMap()
            {
                Map(p => p.SubValue1.StringValue);
                Map(p => p.SubValue1.SplitStringValue);
                Map(p => p.SubValue2.IntValue);
                Map(p => p.SubValue2.SubValue.SubInt);
                Map(p => p.SubValue2.SubValue.SubString);
            }
        }
    }
}

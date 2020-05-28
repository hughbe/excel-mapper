using Xunit;

namespace ExcelMapper.Tests
{
    public class MapCustomColumnNameTests
    {
        [Fact]
        public void ReadRows_AutoMappedCustomIndexProperty_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            CustomIndexPropertyClass row1 = sheet.ReadRow<CustomIndexPropertyClass>();
            Assert.Equal("a", row1.CustomIndex);
        }

        [Fact]
        public void ReadRows_AutoMappedCustomIndexField_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            CustomIndexFieldClass row1 = sheet.ReadRow<CustomIndexFieldClass>();
            Assert.Equal("a", row1.CustomIndex);
        }

        [Fact]
        public void ReadRows_DefaultMappedCustomIndexProperty_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            importer.Configuration.RegisterClassMap<DefaultCustomIndexPropertyClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            CustomIndexPropertyClass row1 = sheet.ReadRow<CustomIndexPropertyClass>();
            Assert.Equal("a", row1.CustomIndex);
        }

        [Fact]
        public void ReadRows_DefaultMappedCustomIndexField_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            importer.Configuration.RegisterClassMap<DefaultCustomIndexFieldClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            CustomIndexFieldClass row1 = sheet.ReadRow<CustomIndexFieldClass>();
            Assert.Equal("a", row1.CustomIndex);
        }

        [Fact]
        public void ReadRows_CustomMappedCustomIndexProperty_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            importer.Configuration.RegisterClassMap<CustomCustomIndexPropertyClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            CustomIndexPropertyClass row1 = sheet.ReadRow<CustomIndexPropertyClass>();
            Assert.Equal("1", row1.CustomIndex);
        }

        [Fact]
        public void ReadRows_CustomMappedCustomIndexField_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            importer.Configuration.RegisterClassMap<CustomCustomIndexFieldClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            CustomIndexFieldClass row1 = sheet.ReadRow<CustomIndexFieldClass>();
            Assert.Equal("1", row1.CustomIndex);
        }

        private class CustomIndexPropertyClass
        {
            [ExcelColumnIndex(1)]
            public string CustomIndex { get; set; }
        }

        private class DefaultCustomIndexPropertyClassMap : ExcelClassMap<CustomIndexPropertyClass>
        {
            public DefaultCustomIndexPropertyClassMap()
            {
                Map(p => p.CustomIndex);
            }
        }

        private class CustomCustomIndexPropertyClassMap : ExcelClassMap<CustomIndexPropertyClass>
        {
            public CustomCustomIndexPropertyClassMap()
            {
                Map(p => p.CustomIndex)
                    .WithColumnName("Int Value");
            }
        }

#pragma warning disable CS0649
        private class CustomIndexFieldClass
        {
            [ExcelColumnIndex(1)]
            public string CustomIndex { get; set; }
        }
#pragma warning restore CS0649

        private class DefaultCustomIndexFieldClassMap : ExcelClassMap<CustomIndexFieldClass>
        {
            public DefaultCustomIndexFieldClassMap()
            {
                Map(p => p.CustomIndex);
            }
        }

        private class CustomCustomIndexFieldClassMap : ExcelClassMap<CustomIndexFieldClass>
        {
            public CustomCustomIndexFieldClassMap()
            {
                Map(p => p.CustomIndex)
                    .WithColumnName("Int Value");
            }
        }

        [Fact]
        public void ReadRows_AutoMappedCustomIndexEnumProperty_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            CustomIndexEnumPropertyClass row1 = sheet.ReadRow<CustomIndexEnumPropertyClass>();
            Assert.Equal(CustomEnum.a, row1.CustomIndex);

            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomIndexEnumPropertyClass>());
        }

        [Fact]
        public void ReadRows_AutoMappedCustomIndexEnumField_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            CustomIndexEnumFieldClass row1 = sheet.ReadRow<CustomIndexEnumFieldClass>();
            Assert.Equal(CustomEnum.a, row1.CustomIndex);

            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomIndexEnumFieldClass>());
        }

        [Fact]
        public void ReadRows_DefaultMappedCustomIndexEnumProperty_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            importer.Configuration.RegisterClassMap<DefaultCustomIndexEnumPropertyClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            CustomIndexEnumPropertyClass row1 = sheet.ReadRow<CustomIndexEnumPropertyClass>();
            Assert.Equal(CustomEnum.a, row1.CustomIndex);

            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomIndexEnumPropertyClass>());
        }

        [Fact]
        public void ReadRows_DefaultMappedCustomIndexEnumField_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            importer.Configuration.RegisterClassMap<DefaultCustomIndexEnumFieldClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            CustomIndexEnumFieldClass row1 = sheet.ReadRow<CustomIndexEnumFieldClass>();
            Assert.Equal(CustomEnum.a, row1.CustomIndex);

            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomIndexEnumFieldClass>());
        }

        [Fact]
        public void ReadRows_CustomMappedCustomIndexEnumProperty_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            importer.Configuration.RegisterClassMap<CustomCustomIndexEnumPropertyClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            CustomIndexEnumPropertyClass row1 = sheet.ReadRow<CustomIndexEnumPropertyClass>();
            Assert.Equal(CustomEnum.a, row1.CustomIndex);

            CustomIndexEnumPropertyClass row2 = sheet.ReadRow<CustomIndexEnumPropertyClass>();
            Assert.Equal(CustomEnum.B, row2.CustomIndex);
        }

        [Fact]
        public void ReadRows_CustomMappedCustomIndexEnumField_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            importer.Configuration.RegisterClassMap<CustomCustomIndexEnumFieldClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            CustomIndexEnumFieldClass row1 = sheet.ReadRow<CustomIndexEnumFieldClass>();
            Assert.Equal(CustomEnum.a, row1.CustomIndex);

            CustomIndexEnumFieldClass row2 = sheet.ReadRow<CustomIndexEnumFieldClass>();
            Assert.Equal(CustomEnum.B, row2.CustomIndex);
        }

        [Fact]
        public void ReadRows_AutoMappedCustomIndexNullableEnumProperty_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            CustomIndexNullableEnumPropertyClass row1 = sheet.ReadRow<CustomIndexNullableEnumPropertyClass>();
            Assert.Equal(CustomEnum.a, row1.CustomIndex);

            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomIndexNullableEnumPropertyClass>());
        }


        [Fact]
        public void ReadRows_AutoMappedCustomIndexNullableEnumField_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            CustomIndexNullableEnumFieldClass row1 = sheet.ReadRow<CustomIndexNullableEnumFieldClass>();
            Assert.Equal(CustomEnum.a, row1.CustomIndex);

            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomIndexNullableEnumFieldClass>());
        }

        [Fact]
        public void ReadRows_DefaultMappedCustomIndexNullableEnumProperty_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            importer.Configuration.RegisterClassMap<DefaultCustomIndexNullableEnumPropertyClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            CustomIndexNullableEnumPropertyClass row1 = sheet.ReadRow<CustomIndexNullableEnumPropertyClass>();
            Assert.Equal(CustomEnum.a, row1.CustomIndex);

            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomIndexNullableEnumPropertyClass>());
        }

        [Fact]
        public void ReadRows_DefaultMappedCustomIndexNullableEnumField_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            importer.Configuration.RegisterClassMap<DefaultCustomIndexNullableEnumFieldClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            CustomIndexNullableEnumFieldClass row1 = sheet.ReadRow<CustomIndexNullableEnumFieldClass>();
            Assert.Equal(CustomEnum.a, row1.CustomIndex);

            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomIndexNullableEnumFieldClass>());
        }

        [Fact]
        public void ReadRows_CustomMappedCustomIndexNullableEnumProperty_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            importer.Configuration.RegisterClassMap<CustomCustomIndexNullableEnumPropertyClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            CustomIndexNullableEnumPropertyClass row1 = sheet.ReadRow<CustomIndexNullableEnumPropertyClass>();
            Assert.Equal(CustomEnum.a, row1.CustomIndex);

            CustomIndexNullableEnumPropertyClass row2 = sheet.ReadRow<CustomIndexNullableEnumPropertyClass>();
            Assert.Equal(CustomEnum.B, row2.CustomIndex);
        }

        [Fact]
        public void ReadRows_CustomMappedCustomIndexNullableEnumField_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            importer.Configuration.RegisterClassMap<CustomCustomIndexNullableEnumFieldClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            CustomIndexNullableEnumFieldClass row1 = sheet.ReadRow<CustomIndexNullableEnumFieldClass>();
            Assert.Equal(CustomEnum.a, row1.CustomIndex);

            CustomIndexNullableEnumFieldClass row2 = sheet.ReadRow<CustomIndexNullableEnumFieldClass>();
            Assert.Equal(CustomEnum.B, row2.CustomIndex);
        }

        private class CustomIndexEnumPropertyClass
        {
            [ExcelColumnIndex(1)]
            public CustomEnum CustomIndex { get; set; }
        }

        private class DefaultCustomIndexEnumPropertyClassMap : ExcelClassMap<CustomIndexEnumPropertyClass>
        {
            public DefaultCustomIndexEnumPropertyClassMap()
            {
                Map(p => p.CustomIndex);
            }
        }

        private class CustomCustomIndexEnumPropertyClassMap : ExcelClassMap<CustomIndexEnumPropertyClass>
        {
            public CustomCustomIndexEnumPropertyClassMap()
            {
                Map(p => p.CustomIndex, ignoreCase: true);
            }
        }

        private class CustomIndexNullableEnumPropertyClass
        {
            [ExcelColumnIndex(1)]
            public CustomEnum? CustomIndex { get; set; }
        }

        private class DefaultCustomIndexNullableEnumPropertyClassMap : ExcelClassMap<CustomIndexNullableEnumPropertyClass>
        {
            public DefaultCustomIndexNullableEnumPropertyClassMap()
            {
                Map(p => p.CustomIndex);
            }
        }

        private class CustomCustomIndexNullableEnumPropertyClassMap : ExcelClassMap<CustomIndexNullableEnumPropertyClass>
        {
            public CustomCustomIndexNullableEnumPropertyClassMap()
            {
                Map(p => p.CustomIndex, ignoreCase: true);
            }
        }
#pragma warning disable CS0649
        private class CustomIndexEnumFieldClass
        {
            [ExcelColumnIndex(1)]
            public CustomEnum CustomIndex { get; set; }
        }
#pragma warning restore CS0649

        private class DefaultCustomIndexEnumFieldClassMap : ExcelClassMap<CustomIndexEnumFieldClass>
        {
            public DefaultCustomIndexEnumFieldClassMap()
            {
                Map(p => p.CustomIndex);
            }
        }

        private class CustomCustomIndexEnumFieldClassMap : ExcelClassMap<CustomIndexEnumFieldClass>
        {
            public CustomCustomIndexEnumFieldClassMap()
            {
                Map(p => p.CustomIndex, ignoreCase: true);
            }
        }

#pragma warning disable CS0649
        private class CustomIndexNullableEnumFieldClass
        {
            [ExcelColumnIndex(1)]
            public CustomEnum? CustomIndex { get; set; }
        }
#pragma warning restore CS0649

        private class DefaultCustomIndexNullableEnumFieldClassMap : ExcelClassMap<CustomIndexNullableEnumFieldClass>
        {
            public DefaultCustomIndexNullableEnumFieldClassMap()
            {
                Map(p => p.CustomIndex);
            }
        }

        private class CustomCustomIndexNullableEnumFieldClassMap : ExcelClassMap<CustomIndexNullableEnumFieldClass>
        {
            public CustomCustomIndexNullableEnumFieldClassMap()
            {
                Map(p => p.CustomIndex, ignoreCase: true);
            }
        }

        private enum CustomEnum
        {
            a,
            B
        }

        [Fact]
        public void ReadRows_AutoMappedCustomIndexEnumerableProperty_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            CustomIndexEnumerablePropertyClass row1 = sheet.ReadRow<CustomIndexEnumerablePropertyClass>();
            Assert.Equal(new string[] { "1", "2", "3" }, row1.CustomValue);

            CustomIndexEnumerablePropertyClass row2 = sheet.ReadRow<CustomIndexEnumerablePropertyClass>();
            Assert.Equal(new string[] { "1", null, "2" }, row2.CustomValue);

            CustomIndexEnumerablePropertyClass row3 = sheet.ReadRow<CustomIndexEnumerablePropertyClass>();
            Assert.Equal(new string[] { "1" }, row3.CustomValue);

            CustomIndexEnumerablePropertyClass row4 = sheet.ReadRow<CustomIndexEnumerablePropertyClass>();
            Assert.Empty(row4.CustomValue);

            CustomIndexEnumerablePropertyClass row5 = sheet.ReadRow<CustomIndexEnumerablePropertyClass>();
            Assert.Equal(new string[] { "Invalid" }, row5.CustomValue);
        }

        [Fact]
        public void ReadRows_AutoMappedCustomIndexEnumerableField_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            CustomIndexEnumerableFieldClass row1 = sheet.ReadRow<CustomIndexEnumerableFieldClass>();
            Assert.Equal(new string[] { "1", "2", "3" }, row1.CustomValue);

            CustomIndexEnumerableFieldClass row2 = sheet.ReadRow<CustomIndexEnumerableFieldClass>();
            Assert.Equal(new string[] { "1", null, "2" }, row2.CustomValue);

            CustomIndexEnumerableFieldClass row3 = sheet.ReadRow<CustomIndexEnumerableFieldClass>();
            Assert.Equal(new string[] { "1" }, row3.CustomValue);

            CustomIndexEnumerableFieldClass row4 = sheet.ReadRow<CustomIndexEnumerableFieldClass>();
            Assert.Empty(row4.CustomValue);

            CustomIndexEnumerableFieldClass row5 = sheet.ReadRow<CustomIndexEnumerableFieldClass>();
            Assert.Equal(new string[] { "Invalid" }, row5.CustomValue);
        }

        [Fact]
        public void ReadRows_DefaultMappedCustomIndexEnumerableProperty_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
            importer.Configuration.RegisterClassMap<DefaultCustomIndexEnumerablePropertyClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            CustomIndexEnumerablePropertyClass row1 = sheet.ReadRow<CustomIndexEnumerablePropertyClass>();
            Assert.Equal(new string[] { "1", "2", "3" }, row1.CustomValue);

            CustomIndexEnumerablePropertyClass row2 = sheet.ReadRow<CustomIndexEnumerablePropertyClass>();
            Assert.Equal(new string[] { "1", null, "2" }, row2.CustomValue);

            CustomIndexEnumerablePropertyClass row3 = sheet.ReadRow<CustomIndexEnumerablePropertyClass>();
            Assert.Equal(new string[] { "1" }, row3.CustomValue);

            CustomIndexEnumerablePropertyClass row4 = sheet.ReadRow<CustomIndexEnumerablePropertyClass>();
            Assert.Empty(row4.CustomValue);

            CustomIndexEnumerablePropertyClass row5 = sheet.ReadRow<CustomIndexEnumerablePropertyClass>();
            Assert.Equal(new string[] { "Invalid" }, row5.CustomValue);
        }

        [Fact]
        public void ReadRows_DefaultMappedCustomIndexEnumerableField_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
            importer.Configuration.RegisterClassMap<DefaultCustomIndexEnumerableFieldClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            CustomIndexEnumerableFieldClass row1 = sheet.ReadRow<CustomIndexEnumerableFieldClass>();
            Assert.Equal(new string[] { "1", "2", "3" }, row1.CustomValue);

            CustomIndexEnumerableFieldClass row2 = sheet.ReadRow<CustomIndexEnumerableFieldClass>();
            Assert.Equal(new string[] { "1", null, "2" }, row2.CustomValue);

            CustomIndexEnumerableFieldClass row3 = sheet.ReadRow<CustomIndexEnumerableFieldClass>();
            Assert.Equal(new string[] { "1" }, row3.CustomValue);

            CustomIndexEnumerableFieldClass row4 = sheet.ReadRow<CustomIndexEnumerableFieldClass>();
            Assert.Empty(row4.CustomValue);

            CustomIndexEnumerableFieldClass row5 = sheet.ReadRow<CustomIndexEnumerableFieldClass>();
            Assert.Equal(new string[] { "Invalid" }, row5.CustomValue);
        }

        private class CustomIndexEnumerablePropertyClass
        {
            [ExcelColumnIndex(0)]
            public object[] CustomValue { get; set; }
        }

        private class DefaultCustomIndexEnumerablePropertyClassMap : ExcelClassMap<CustomIndexEnumerablePropertyClass>
        {
            public DefaultCustomIndexEnumerablePropertyClassMap()
            {
                Map(p => p.CustomValue);
            }
        }

#pragma warning disable CS0649
        private class CustomIndexEnumerableFieldClass
        {
            [ExcelColumnIndex(0)]
            public object[] CustomValue { get; set; }
        }
#pragma warning restore CS0649

        private class DefaultCustomIndexEnumerableFieldClassMap : ExcelClassMap<CustomIndexEnumerableFieldClass>
        {
            public DefaultCustomIndexEnumerableFieldClassMap()
            {
                Map(p => p.CustomValue);
            }
        }
    }
}

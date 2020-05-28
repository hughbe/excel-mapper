using Xunit;

namespace ExcelMapper.Tests
{
    public class MapCustomColumnIndexTests
    {
        [Fact]
        public void ReadRows_AutoMappedCustomNameProperty_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            CustomNamePropertyClass row1 = sheet.ReadRow<CustomNamePropertyClass>();
            Assert.Equal("a", row1.CustomName);
        }

        [Fact]
        public void ReadRows_AutoMappedCustomNameField_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            CustomNameFieldClass row1 = sheet.ReadRow<CustomNameFieldClass>();
            Assert.Equal("a", row1.CustomName);
        }

        [Fact]
        public void ReadRows_DefaultMappedCustomNameProperty_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            importer.Configuration.RegisterClassMap<DefaultCustomNamePropertyClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            CustomNamePropertyClass row1 = sheet.ReadRow<CustomNamePropertyClass>();
            Assert.Equal("a", row1.CustomName);
        }

        [Fact]
        public void ReadRows_DefaultMappedCustomNameField_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            importer.Configuration.RegisterClassMap<DefaultCustomNameFieldClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            CustomNameFieldClass row1 = sheet.ReadRow<CustomNameFieldClass>();
            Assert.Equal("a", row1.CustomName);
        }

        [Fact]
        public void ReadRows_CustomMappedCustomNameProperty_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            importer.Configuration.RegisterClassMap<CustomCustomNamePropertyClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            CustomNamePropertyClass row1 = sheet.ReadRow<CustomNamePropertyClass>();
            Assert.Equal("1", row1.CustomName);
        }

        [Fact]
        public void ReadRows_CustomMappedCustomNameField_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            importer.Configuration.RegisterClassMap<CustomCustomNameFieldClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            CustomNameFieldClass row1 = sheet.ReadRow<CustomNameFieldClass>();
            Assert.Equal("1", row1.CustomName);
        }

        private class CustomNamePropertyClass
        {
            [ExcelColumnName("StringValue")]
            public string CustomName { get; set; }
        }

        private class DefaultCustomNamePropertyClassMap : ExcelClassMap<CustomNamePropertyClass>
        {
            public DefaultCustomNamePropertyClassMap()
            {
                Map(p => p.CustomName);
            }
        }

        private class CustomCustomNamePropertyClassMap : ExcelClassMap<CustomNamePropertyClass>
        {
            public CustomCustomNamePropertyClassMap()
            {
                Map(p => p.CustomName)
                    .WithColumnName("Int Value");
            }
        }

#pragma warning disable CS0649
        private class CustomNameFieldClass
        {
            [ExcelColumnName("StringValue")]
            public string CustomName { get; set; }
        }
#pragma warning restore CS0649

        private class DefaultCustomNameFieldClassMap : ExcelClassMap<CustomNameFieldClass>
        {
            public DefaultCustomNameFieldClassMap()
            {
                Map(p => p.CustomName);
            }
        }

        private class CustomCustomNameFieldClassMap : ExcelClassMap<CustomNameFieldClass>
        {
            public CustomCustomNameFieldClassMap()
            {
                Map(p => p.CustomName)
                    .WithColumnName("Int Value");
            }
        }

        [Fact]
        public void ReadRows_AutoMappedCustomNameEnumProperty_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            CustomNameEnumPropertyClass row1 = sheet.ReadRow<CustomNameEnumPropertyClass>();
            Assert.Equal(CustomEnum.a, row1.CustomName);

            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNameEnumPropertyClass>());
        }

        [Fact]
        public void ReadRows_AutoMappedCustomNameEnumField_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            CustomNameEnumFieldClass row1 = sheet.ReadRow<CustomNameEnumFieldClass>();
            Assert.Equal(CustomEnum.a, row1.CustomName);

            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNameEnumFieldClass>());
        }

        [Fact]
        public void ReadRows_DefaultMappedCustomNameEnumProperty_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            importer.Configuration.RegisterClassMap<DefaultCustomNameEnumPropertyClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            CustomNameEnumPropertyClass row1 = sheet.ReadRow<CustomNameEnumPropertyClass>();
            Assert.Equal(CustomEnum.a, row1.CustomName);

            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNameEnumPropertyClass>());
        }

        [Fact]
        public void ReadRows_DefaultMappedCustomNameEnumField_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            importer.Configuration.RegisterClassMap<DefaultCustomNameEnumFieldClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            CustomNameEnumFieldClass row1 = sheet.ReadRow<CustomNameEnumFieldClass>();
            Assert.Equal(CustomEnum.a, row1.CustomName);

            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNameEnumFieldClass>());
        }

        [Fact]
        public void ReadRows_CustomMappedCustomNameEnumProperty_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            importer.Configuration.RegisterClassMap<CustomCustomNameEnumPropertyClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            CustomNameEnumPropertyClass row1 = sheet.ReadRow<CustomNameEnumPropertyClass>();
            Assert.Equal(CustomEnum.a, row1.CustomName);

            CustomNameEnumPropertyClass row2 = sheet.ReadRow<CustomNameEnumPropertyClass>();
            Assert.Equal(CustomEnum.B, row2.CustomName);
        }

        [Fact]
        public void ReadRows_CustomMappedCustomNameEnumField_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            importer.Configuration.RegisterClassMap<CustomCustomNameEnumFieldClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            CustomNameEnumFieldClass row1 = sheet.ReadRow<CustomNameEnumFieldClass>();
            Assert.Equal(CustomEnum.a, row1.CustomName);

            CustomNameEnumFieldClass row2 = sheet.ReadRow<CustomNameEnumFieldClass>();
            Assert.Equal(CustomEnum.B, row2.CustomName);
        }

        [Fact]
        public void ReadRows_AutoMappedCustomNameNullableEnumProperty_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            CustomNameNullableEnumPropertyClass row1 = sheet.ReadRow<CustomNameNullableEnumPropertyClass>();
            Assert.Equal(CustomEnum.a, row1.CustomName);

            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNameNullableEnumPropertyClass>());
        }


        [Fact]
        public void ReadRows_AutoMappedCustomNameNullableEnumField_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            CustomNameNullableEnumFieldClass row1 = sheet.ReadRow<CustomNameNullableEnumFieldClass>();
            Assert.Equal(CustomEnum.a, row1.CustomName);

            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNameNullableEnumFieldClass>());
        }

        [Fact]
        public void ReadRows_DefaultMappedCustomNameNullableEnumProperty_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            importer.Configuration.RegisterClassMap<DefaultCustomNameNullableEnumPropertyClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            CustomNameNullableEnumPropertyClass row1 = sheet.ReadRow<CustomNameNullableEnumPropertyClass>();
            Assert.Equal(CustomEnum.a, row1.CustomName);

            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNameNullableEnumPropertyClass>());
        }

        [Fact]
        public void ReadRows_DefaultMappedCustomNameNullableEnumField_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            importer.Configuration.RegisterClassMap<DefaultCustomNameNullableEnumFieldClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            CustomNameNullableEnumFieldClass row1 = sheet.ReadRow<CustomNameNullableEnumFieldClass>();
            Assert.Equal(CustomEnum.a, row1.CustomName);

            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<CustomNameNullableEnumFieldClass>());
        }

        [Fact]
        public void ReadRows_CustomMappedCustomNameNullableEnumProperty_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            importer.Configuration.RegisterClassMap<CustomCustomNameNullableEnumPropertyClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            CustomNameNullableEnumPropertyClass row1 = sheet.ReadRow<CustomNameNullableEnumPropertyClass>();
            Assert.Equal(CustomEnum.a, row1.CustomName);

            CustomNameNullableEnumPropertyClass row2 = sheet.ReadRow<CustomNameNullableEnumPropertyClass>();
            Assert.Equal(CustomEnum.B, row2.CustomName);
        }

        [Fact]
        public void ReadRows_CustomMappedCustomNameNullableEnumField_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            importer.Configuration.RegisterClassMap<CustomCustomNameNullableEnumFieldClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            CustomNameNullableEnumFieldClass row1 = sheet.ReadRow<CustomNameNullableEnumFieldClass>();
            Assert.Equal(CustomEnum.a, row1.CustomName);

            CustomNameNullableEnumFieldClass row2 = sheet.ReadRow<CustomNameNullableEnumFieldClass>();
            Assert.Equal(CustomEnum.B, row2.CustomName);
        }

        private class CustomNameEnumPropertyClass
        {
            [ExcelColumnName("StringValue")]
            public CustomEnum CustomName { get; set; }
        }

        private class DefaultCustomNameEnumPropertyClassMap : ExcelClassMap<CustomNameEnumPropertyClass>
        {
            public DefaultCustomNameEnumPropertyClassMap()
            {
                Map(p => p.CustomName);
            }
        }

        private class CustomCustomNameEnumPropertyClassMap : ExcelClassMap<CustomNameEnumPropertyClass>
        {
            public CustomCustomNameEnumPropertyClassMap()
            {
                Map(p => p.CustomName, ignoreCase: true);
            }
        }

        private class CustomNameNullableEnumPropertyClass
        {
            [ExcelColumnName("StringValue")]
            public CustomEnum? CustomName { get; set; }
        }

        private class DefaultCustomNameNullableEnumPropertyClassMap : ExcelClassMap<CustomNameNullableEnumPropertyClass>
        {
            public DefaultCustomNameNullableEnumPropertyClassMap()
            {
                Map(p => p.CustomName);
            }
        }

        private class CustomCustomNameNullableEnumPropertyClassMap : ExcelClassMap<CustomNameNullableEnumPropertyClass>
        {
            public CustomCustomNameNullableEnumPropertyClassMap()
            {
                Map(p => p.CustomName, ignoreCase: true);
            }
        }
#pragma warning disable CS0649
        private class CustomNameEnumFieldClass
        {
            [ExcelColumnName("StringValue")]
            public CustomEnum CustomName { get; set; }
        }
#pragma warning restore CS0649

        private class DefaultCustomNameEnumFieldClassMap : ExcelClassMap<CustomNameEnumFieldClass>
        {
            public DefaultCustomNameEnumFieldClassMap()
            {
                Map(p => p.CustomName);
            }
        }

        private class CustomCustomNameEnumFieldClassMap : ExcelClassMap<CustomNameEnumFieldClass>
        {
            public CustomCustomNameEnumFieldClassMap()
            {
                Map(p => p.CustomName, ignoreCase: true);
            }
        }

#pragma warning disable CS0649
        private class CustomNameNullableEnumFieldClass
        {
            [ExcelColumnName("StringValue")]
            public CustomEnum? CustomName { get; set; }
        }
#pragma warning restore CS0649

        private class DefaultCustomNameNullableEnumFieldClassMap : ExcelClassMap<CustomNameNullableEnumFieldClass>
        {
            public DefaultCustomNameNullableEnumFieldClassMap()
            {
                Map(p => p.CustomName);
            }
        }

        private class CustomCustomNameNullableEnumFieldClassMap : ExcelClassMap<CustomNameNullableEnumFieldClass>
        {
            public CustomCustomNameNullableEnumFieldClassMap()
            {
                Map(p => p.CustomName, ignoreCase: true);
            }
        }

        private enum CustomEnum
        {
            a,
            B
        }

        [Fact]
        public void ReadRows_AutoMappedCustomNameEnumerableProperty_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            CustomNameEnumerablePropertyClass row1 = sheet.ReadRow<CustomNameEnumerablePropertyClass>();
            Assert.Equal(new string[] { "1", "2", "3" }, row1.CustomValue);

            CustomNameEnumerablePropertyClass row2 = sheet.ReadRow<CustomNameEnumerablePropertyClass>();
            Assert.Equal(new string[] { "1", null, "2" }, row2.CustomValue);

            CustomNameEnumerablePropertyClass row3 = sheet.ReadRow<CustomNameEnumerablePropertyClass>();
            Assert.Equal(new string[] { "1" }, row3.CustomValue);

            CustomNameEnumerablePropertyClass row4 = sheet.ReadRow<CustomNameEnumerablePropertyClass>();
            Assert.Empty(row4.CustomValue);

            CustomNameEnumerablePropertyClass row5 = sheet.ReadRow<CustomNameEnumerablePropertyClass>();
            Assert.Equal(new string[] { "Invalid" }, row5.CustomValue);
        }

        [Fact]
        public void ReadRows_AutoMappedCustomNameEnumerableField_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("SplitWithComma.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            CustomNameEnumerableFieldClass row1 = sheet.ReadRow<CustomNameEnumerableFieldClass>();
            Assert.Equal(new string[] { "1", "2", "3" }, row1.CustomValue);

            CustomNameEnumerableFieldClass row2 = sheet.ReadRow<CustomNameEnumerableFieldClass>();
            Assert.Equal(new string[] { "1", null, "2" }, row2.CustomValue);

            CustomNameEnumerableFieldClass row3 = sheet.ReadRow<CustomNameEnumerableFieldClass>();
            Assert.Equal(new string[] { "1" }, row3.CustomValue);

            CustomNameEnumerableFieldClass row4 = sheet.ReadRow<CustomNameEnumerableFieldClass>();
            Assert.Empty(row4.CustomValue);

            CustomNameEnumerableFieldClass row5 = sheet.ReadRow<CustomNameEnumerableFieldClass>();
            Assert.Equal(new string[] { "Invalid" }, row5.CustomValue);
        }

        [Fact]
        public void ReadRows_DefaultMappedCustomNameEnumerableProperty_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
            importer.Configuration.RegisterClassMap<DefaultCustomNameEnumerablePropertyClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            CustomNameEnumerablePropertyClass row1 = sheet.ReadRow<CustomNameEnumerablePropertyClass>();
            Assert.Equal(new string[] { "1", "2", "3" }, row1.CustomValue);

            CustomNameEnumerablePropertyClass row2 = sheet.ReadRow<CustomNameEnumerablePropertyClass>();
            Assert.Equal(new string[] { "1", null, "2" }, row2.CustomValue);

            CustomNameEnumerablePropertyClass row3 = sheet.ReadRow<CustomNameEnumerablePropertyClass>();
            Assert.Equal(new string[] { "1" }, row3.CustomValue);

            CustomNameEnumerablePropertyClass row4 = sheet.ReadRow<CustomNameEnumerablePropertyClass>();
            Assert.Empty(row4.CustomValue);

            CustomNameEnumerablePropertyClass row5 = sheet.ReadRow<CustomNameEnumerablePropertyClass>();
            Assert.Equal(new string[] { "Invalid" }, row5.CustomValue);
        }

        [Fact]
        public void ReadRows_DefaultMappedCustomNameEnumerableField_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("SplitWithComma.xlsx");
            importer.Configuration.RegisterClassMap<DefaultCustomNameEnumerableFieldClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            CustomNameEnumerableFieldClass row1 = sheet.ReadRow<CustomNameEnumerableFieldClass>();
            Assert.Equal(new string[] { "1", "2", "3" }, row1.CustomValue);

            CustomNameEnumerableFieldClass row2 = sheet.ReadRow<CustomNameEnumerableFieldClass>();
            Assert.Equal(new string[] { "1", null, "2" }, row2.CustomValue);

            CustomNameEnumerableFieldClass row3 = sheet.ReadRow<CustomNameEnumerableFieldClass>();
            Assert.Equal(new string[] { "1" }, row3.CustomValue);

            CustomNameEnumerableFieldClass row4 = sheet.ReadRow<CustomNameEnumerableFieldClass>();
            Assert.Empty(row4.CustomValue);

            CustomNameEnumerableFieldClass row5 = sheet.ReadRow<CustomNameEnumerableFieldClass>();
            Assert.Equal(new string[] { "Invalid" }, row5.CustomValue);
        }

        private class CustomNameEnumerablePropertyClass
        {
            [ExcelColumnName("Value")]
            public object[] CustomValue { get; set; }
        }

        private class DefaultCustomNameEnumerablePropertyClassMap : ExcelClassMap<CustomNameEnumerablePropertyClass>
        {
            public DefaultCustomNameEnumerablePropertyClassMap()
            {
                Map(p => p.CustomValue);
            }
        }

#pragma warning disable CS0649
        private class CustomNameEnumerableFieldClass
        {
            [ExcelColumnName("Value")]
            public object[] CustomValue { get; set; }
        }
#pragma warning restore CS0649

        private class DefaultCustomNameEnumerableFieldClassMap : ExcelClassMap<CustomNameEnumerableFieldClass>
        {
            public DefaultCustomNameEnumerableFieldClassMap()
            {
                Map(p => p.CustomValue);
            }
        }
    }
}

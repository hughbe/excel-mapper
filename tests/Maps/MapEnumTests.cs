using Xunit;

namespace ExcelMapper.Tests
{
    public class MapEnumTests
    {
        [Fact]
        public void ReadRow_AutoMappedEnum_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Enums.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            EnumClass row1 = sheet.ReadRow<EnumClass>();
            Assert.Equal(TestEnum.Member, row1.Value);

            // Different case cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableEnumClass>());

            // Empty cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EnumClass>());

            // Invalid cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EnumClass>());
        }

        [Fact]
        public void ReadRow_AutoMappedNullableEnum_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Enums.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            NullableEnumClass row1 = sheet.ReadRow<NullableEnumClass>();
            Assert.Equal(TestEnum.Member, row1.Value);

            // Different case cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableEnumClass>());

            // Empty cell value.
            NullableEnumClass row3 = sheet.ReadRow<NullableEnumClass>();
            Assert.Null(row3.Value);

            // Invalid cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableEnumClass>());
        }

        [Fact]
        public void ReadRow_DefaultMappedEnum_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Enums.xlsx");
            importer.Configuration.RegisterClassMap<DefaultEnumClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            EnumClass row1 = sheet.ReadRow<EnumClass>();
            Assert.Equal(TestEnum.Member, row1.Value);

            // Different case cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableEnumClass>());

            // Empty cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EnumClass>());

            // Invalid cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EnumClass>());
        }

        [Fact]
        public void ReadRow_DefaultMappedNullableEnum_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Enums.xlsx");
            importer.Configuration.RegisterClassMap<DefaultNullableCustomEnumClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            NullableEnumClass row1 = sheet.ReadRow<NullableEnumClass>();
            Assert.Equal(TestEnum.Member, row1.Value);

            // Different case cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableEnumClass>());

            // Empty cell value.
            NullableEnumClass row3 = sheet.ReadRow<NullableEnumClass>();
            Assert.Null(row3.Value);

            // Invalid cell value.
            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableEnumClass>());
        }

        [Fact]
        public void ReadRow_CustomMappedEnum_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Enums.xlsx");
            importer.Configuration.RegisterClassMap<CustomEnumClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            EnumClass row1 = sheet.ReadRow<EnumClass>();
            Assert.Equal(TestEnum.Member, row1.Value);

            // Different case cell value.
            EnumClass row2 = sheet.ReadRow<EnumClass>();
            Assert.Equal(TestEnum.Invalid, row2.Value);

            // Empty cell value.
            EnumClass row3 = sheet.ReadRow<EnumClass>();
            Assert.Equal(TestEnum.Empty, row3.Value);

            // Invalid cell value.
            EnumClass row4 = sheet.ReadRow<EnumClass>();
            Assert.Equal(TestEnum.Invalid, row4.Value);
        }

        [Fact]
        public void ReadRow_IgnoreCaseEnum_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Enums.xlsx");
            importer.Configuration.RegisterClassMap<IgnoreCaseEnumClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            EnumClass row1 = sheet.ReadRow<EnumClass>();
            Assert.Equal(TestEnum.Member, row1.Value);

            // Different case cell value.
            EnumClass row2 = sheet.ReadRow<EnumClass>();
            Assert.Equal(TestEnum.Member, row2.Value);

            // Empty cell value.
            EnumClass row3 = sheet.ReadRow<EnumClass>();
            Assert.Equal(TestEnum.Empty, row3.Value);

            // Invalid cell value.
            EnumClass row4 = sheet.ReadRow<EnumClass>();
            Assert.Equal(TestEnum.Invalid, row4.Value);
        }

        [Fact]
        public void ReadRow_NullableCustomMappedEnum_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Enums.xlsx");
            importer.Configuration.RegisterClassMap<CustomNullableCustomEnumClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            NullableEnumClass row1 = sheet.ReadRow<NullableEnumClass>();
            Assert.Equal(TestEnum.Member, row1.Value);

            // Different case cell value.
            NullableEnumClass row2 = sheet.ReadRow<NullableEnumClass>();
            Assert.Equal(TestEnum.Invalid, row2.Value);

            // Empty cell value.
            NullableEnumClass row3 = sheet.ReadRow<NullableEnumClass>();
            Assert.Equal(TestEnum.Empty, row3.Value);

            // Invalid cell value.
            NullableEnumClass row4 = sheet.ReadRow<NullableEnumClass>();
            Assert.Equal(TestEnum.Invalid, row4.Value);
        }

        [Fact]
        public void ReadRow_NullableIgnoreCaseEnum_ReturnsExpected()
        {
            using var importer = Helpers.GetImporter("Enums.xlsx");
            importer.Configuration.RegisterClassMap<IgnoreCaseNullableCustomEnumClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            // Valid cell value.
            NullableEnumClass row1 = sheet.ReadRow<NullableEnumClass>();
            Assert.Equal(TestEnum.Member, row1.Value);

            // Different case cell value.
            NullableEnumClass row2 = sheet.ReadRow<NullableEnumClass>();
            Assert.Equal(TestEnum.Member, row2.Value);

            // Empty cell value.
            NullableEnumClass row3 = sheet.ReadRow<NullableEnumClass>();
            Assert.Equal(TestEnum.Empty, row3.Value);

            // Invalid cell value.
            NullableEnumClass row4 = sheet.ReadRow<NullableEnumClass>();
            Assert.Equal(TestEnum.Invalid, row4.Value);
        }

        private enum TestEnum
        {
            Member,
            Empty,
            Invalid
        }

        private class EnumClass
        {
            public TestEnum Value { get; set; }
        }

        private class DefaultEnumClassMap : ExcelClassMap<EnumClass>
        {
            public DefaultEnumClassMap()
            {
                Map(u => u.Value);
            }
        }

        private class CustomEnumClassMap : ExcelClassMap<EnumClass>
        {
            public CustomEnumClassMap()
            {
                Map(u => u.Value)
                    .WithEmptyFallback(TestEnum.Empty)
                    .WithInvalidFallback(TestEnum.Invalid);
            }
        }

        private class IgnoreCaseEnumClassMap : ExcelClassMap<EnumClass>
        {
            public IgnoreCaseEnumClassMap()
            {
                Map(u => u.Value, ignoreCase: true)
                    .WithEmptyFallback(TestEnum.Empty)
                    .WithInvalidFallback(TestEnum.Invalid);
            }
        }

        private class NullableEnumClass
        {
            public TestEnum? Value { get; set; }
        }

        private class DefaultNullableCustomEnumClassMap : ExcelClassMap<NullableEnumClass>
        {
            public DefaultNullableCustomEnumClassMap()
            {
                Map(u => u.Value);
            }
        }

        private class CustomNullableCustomEnumClassMap : ExcelClassMap<NullableEnumClass>
        {
            public CustomNullableCustomEnumClassMap()
            {
                Map(u => u.Value)
                    .WithEmptyFallback(TestEnum.Empty)
                    .WithInvalidFallback(TestEnum.Invalid);
            }
        }

        private class IgnoreCaseNullableCustomEnumClassMap : ExcelClassMap<NullableEnumClass>
        {
            public IgnoreCaseNullableCustomEnumClassMap()
            {
                Map(u => u.Value, ignoreCase: true)
                    .WithEmptyFallback(TestEnum.Empty)
                    .WithInvalidFallback(TestEnum.Invalid);
            }
        }
    }
}

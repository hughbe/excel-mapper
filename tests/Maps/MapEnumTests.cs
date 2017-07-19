using Xunit;

namespace ExcelMapper.Tests
{
    public class MapEnumTests
    {
        [Fact]
        public void ReadRow_AutoMappedEnum_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("Enums.xlsx"))
            {
                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                // Valid cell value.
                EnumValue row1 = sheet.ReadRow<EnumValue>();
                Assert.Equal(TestEnum.Member, row1.Value);

                // Empty cell value.
                Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EnumValue>());

                // Invalid cell value.
                Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<EnumValue>());
            }
        }

        [Fact]
        public void ReadRow_AutoMappedNullableEnum_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("Enums.xlsx"))
            {
                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                // Valid cell value.
                NullableEnumValue row1 = sheet.ReadRow<NullableEnumValue>();
                Assert.Equal(TestEnum.Member, row1.Value);

                // Empty cell value.
                NullableEnumValue row2 = sheet.ReadRow<NullableEnumValue>();
                Assert.Null(row2.Value);

                // Invalid cell value.
                Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableEnumValue>());
            }
        }

        [Fact]
        public void ReadRow_EnumWithCustomFallback_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("Enums.xlsx"))
            {
                importer.Configuration.RegisterClassMap<EnumValueFallbackMap>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                // Valid cell value.
                EnumValue row1 = sheet.ReadRow<EnumValue>();
                Assert.Equal(TestEnum.Member, row1.Value);

                // Empty cell value.
                EnumValue row2 = sheet.ReadRow<EnumValue>();
                Assert.Equal(TestEnum.Empty, row2.Value);

                // Invalid cell value.
                EnumValue row3 = sheet.ReadRow<EnumValue>();
                Assert.Equal(TestEnum.Invalid, row3.Value);
            }
        }

        [Fact]
        public void ReadRow_NullableEnumWithCustomFallback_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("Enums.xlsx"))
            {
                importer.Configuration.RegisterClassMap<NullableEnumValueFallbackMap>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                // Valid cell value.
                NullableEnumValue row1 = sheet.ReadRow<NullableEnumValue>();
                Assert.Equal(TestEnum.Member, row1.Value);

                // Empty cell value.
                NullableEnumValue row2 = sheet.ReadRow<NullableEnumValue>();
                Assert.Equal(TestEnum.Empty, row2.Value);

                // Invalid cell value.
                NullableEnumValue row3 = sheet.ReadRow<NullableEnumValue>();
                Assert.Equal(TestEnum.Invalid, row3.Value);
            }
        }

        private enum TestEnum
        {
            Member,
            Empty,
            Invalid
        }

        private class EnumValue
        {
            public TestEnum Value { get; set; }
        }

        private class NullableEnumValue
        {
            public TestEnum? Value { get; set; }
        }

        private class EnumValueFallbackMap : ExcelClassMap<EnumValue>
        {
            public EnumValueFallbackMap()
            {
                Map(u => u.Value)
                    .WithEmptyFallback(TestEnum.Empty)
                    .WithInvalidFallback(TestEnum.Invalid);
            }
        }

        private class NullableEnumValueFallbackMap : ExcelClassMap<NullableEnumValue>
        {
            public NullableEnumValueFallbackMap()
            {
                Map(u => u.Value)
                    .WithEmptyFallback(TestEnum.Empty)
                    .WithInvalidFallback(TestEnum.Invalid);
            }
        }
    }
}

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

                // Different case cell value.
                Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableEnumValue>());

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

                // Different case cell value.
                Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NullableEnumValue>());

                // Empty cell value.
                NullableEnumValue row3 = sheet.ReadRow<NullableEnumValue>();
                Assert.Null(row3.Value);

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

                // Different case cell value.
                EnumValue row2 = sheet.ReadRow<EnumValue>();
                Assert.Equal(TestEnum.Invalid, row2.Value);

                // Empty cell value.
                EnumValue row3 = sheet.ReadRow<EnumValue>();
                Assert.Equal(TestEnum.Empty, row3.Value);

                // Invalid cell value.
                EnumValue row4 = sheet.ReadRow<EnumValue>();
                Assert.Equal(TestEnum.Invalid, row4.Value);
            }
        }

        [Fact]
        public void ReadRow_IgnoreCasseEnum_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("Enums.xlsx"))
            {
                importer.Configuration.RegisterClassMap<IgnoreCaseEnumValueFallbackMap>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                // Valid cell value.
                EnumValue row1 = sheet.ReadRow<EnumValue>();
                Assert.Equal(TestEnum.Member, row1.Value);

                // Different case cell value.
                EnumValue row2 = sheet.ReadRow<EnumValue>();
                Assert.Equal(TestEnum.Member, row2.Value);

                // Empty cell value.
                EnumValue row3 = sheet.ReadRow<EnumValue>();
                Assert.Equal(TestEnum.Empty, row3.Value);

                // Invalid cell value.
                EnumValue row4 = sheet.ReadRow<EnumValue>();
                Assert.Equal(TestEnum.Invalid, row4.Value);
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

                // Different case cell value.
                NullableEnumValue row2 = sheet.ReadRow<NullableEnumValue>();
                Assert.Equal(TestEnum.Invalid, row2.Value);

                // Empty cell value.
                NullableEnumValue row3 = sheet.ReadRow<NullableEnumValue>();
                Assert.Equal(TestEnum.Empty, row3.Value);

                // Invalid cell value.
                NullableEnumValue row4 = sheet.ReadRow<NullableEnumValue>();
                Assert.Equal(TestEnum.Invalid, row4.Value);
            }
        }

        [Fact]
        public void ReadRow_NullableIgnoreCaseEnum_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("Enums.xlsx"))
            {
                importer.Configuration.RegisterClassMap<IgnoreCaseNullableEnumValueFallbackMap>();

                ExcelSheet sheet = importer.ReadSheet();
                sheet.ReadHeading();

                // Valid cell value.
                NullableEnumValue row1 = sheet.ReadRow<NullableEnumValue>();
                Assert.Equal(TestEnum.Member, row1.Value);

                // Different case cell value.
                NullableEnumValue row2 = sheet.ReadRow<NullableEnumValue>();
                Assert.Equal(TestEnum.Member, row2.Value);

                // Empty cell value.
                NullableEnumValue row3 = sheet.ReadRow<NullableEnumValue>();
                Assert.Equal(TestEnum.Empty, row3.Value);

                // Invalid cell value.
                NullableEnumValue row4 = sheet.ReadRow<NullableEnumValue>();
                Assert.Equal(TestEnum.Invalid, row4.Value);
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

        private class IgnoreCaseEnumValueFallbackMap : ExcelClassMap<EnumValue>
        {
            public IgnoreCaseEnumValueFallbackMap()
            {
                Map(u => u.Value, ignoreCase: true)
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

        private class IgnoreCaseNullableEnumValueFallbackMap : ExcelClassMap<NullableEnumValue>
        {
            public IgnoreCaseNullableEnumValueFallbackMap()
            {
                Map(u => u.Value, ignoreCase: true)
                    .WithEmptyFallback(TestEnum.Empty)
                    .WithInvalidFallback(TestEnum.Invalid);
            }
        }
    }
}

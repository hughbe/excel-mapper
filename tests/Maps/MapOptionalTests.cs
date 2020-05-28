using Xunit;

namespace ExcelMapper.Tests
{
    public class MapOptionalTests
    {
        [Fact]
        public void ReadRows_AutoMappedPropertyDoesNotExist_ThrowsExcelMappingException()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<MissingColumnPropertyClass>());
        }

        [Fact]
        public void ReadRows_AutoMappedFieldDoesNotExist_ThrowsExcelMappingException()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<MissingColumnFieldClass>());
        }

        [Fact]
        public void ReadRows_DefaultMappedPropertyDoesNotExist_ThrowsExcelMappingException()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            importer.Configuration.RegisterClassMap<DefaultMissingColumnPropertyClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<MissingColumnPropertyClass>());
        }

        [Fact]
        public void ReadRows_DefaultMappedFieldDoesNotExist_ThrowsExcelMappingException()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            importer.Configuration.RegisterClassMap<DefaultMissingColumnFieldClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<MissingColumnFieldClass>());
        }

        [Fact]
        public void ReadRows_CustomMappedPropertyDoesNotExist_Success()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            importer.Configuration.RegisterClassMap<CustomMissingColumnPropertyClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            MissingColumnPropertyClass row1 = sheet.ReadRow<MissingColumnPropertyClass>();
            Assert.Equal(10, row1.NoSuchColumn);
        }

        [Fact]
        public void ReadRows_CustomMappedFieldDoesNotExist_Success()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            importer.Configuration.RegisterClassMap<CustomMissingColumnFieldClassMap>();

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            MissingColumnFieldClass row1 = sheet.ReadRow<MissingColumnFieldClass>();
            Assert.Equal(10, row1.NoSuchColumn);
        }

        private class MissingColumnPropertyClass
        {
            public int NoSuchColumn { get; set; } = 10;
        }

        private class DefaultMissingColumnPropertyClassMap : ExcelClassMap<MissingColumnPropertyClass>
        {
            public DefaultMissingColumnPropertyClassMap()
            {
                Map(p => p.NoSuchColumn);
            }
        }

        private class CustomMissingColumnPropertyClassMap : ExcelClassMap<MissingColumnPropertyClass>
        {
            public CustomMissingColumnPropertyClassMap()
            {
                Map(p => p.NoSuchColumn)
                    .MakeOptional();
            }
        }

#pragma warning disable CS0649
        private class MissingColumnFieldClass
        {
            public int NoSuchColumn = 10;
        }
#pragma warning restore CS0649

        private class DefaultMissingColumnFieldClassMap : ExcelClassMap<MissingColumnFieldClass>
        {
            public DefaultMissingColumnFieldClassMap()
            {
                Map(p => p.NoSuchColumn);
            }
        }

        private class CustomMissingColumnFieldClassMap : ExcelClassMap<MissingColumnFieldClass>
        {
            public CustomMissingColumnFieldClassMap()
            {
                Map(p => p.NoSuchColumn)
                    .MakeOptional();
            }
        }

        [Fact]
        public void ReadRows_IgnoredProperty_DoesNotDeserialize()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            IgnoredColumnPropertyClass row1 = sheet.ReadRow<IgnoredColumnPropertyClass>();
            Assert.Equal("CustomValue", row1.StringValue);
            Assert.Equal("a", row1.MappedValue);
        }

        [Fact]
        public void ReadRows_IgnoredField_DoesNotDeserialize()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            IgnoredColumnFieldClass row1 = sheet.ReadRow<IgnoredColumnFieldClass>();
            Assert.Equal("CustomValue", row1.StringValue);
            Assert.Equal("a", row1.MappedValue);
        }

        [Fact]
        public void ReadRows_IgnoredMissingProperty_DoesNotDeserialize()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            MissingColumnIgnoredPropertyClass row1 = sheet.ReadRow<MissingColumnIgnoredPropertyClass>();
            Assert.Equal(10, row1.NoSuchColumn);
            Assert.Equal("a", row1.MappedValue);
        }

        [Fact]
        public void ReadRows_IgnoredMissingField_DoesNotDeserialize()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            MissingColumnIgnoredFieldClass row1 = sheet.ReadRow<MissingColumnIgnoredFieldClass>();
            Assert.Equal(10, row1.NoSuchColumn);
            Assert.Equal("a", row1.MappedValue);
        }

        private class IgnoredColumnPropertyClass
        {
            [ExcelIgnore]
            public string StringValue { get; set; } = "CustomValue";

            public string MappedValue { get; set; }
        }

        private class MissingColumnIgnoredPropertyClass
        {
            [ExcelIgnore]
            public int NoSuchColumn { get; set; } = 10;

            public string MappedValue { get; set; }
        }

#pragma warning disable CS0649
        private class IgnoredColumnFieldClass
        {
            [ExcelIgnore]
            public string StringValue = "CustomValue";

            public string MappedValue;
        }

        private class MissingColumnIgnoredFieldClass
        {
            [ExcelIgnore]
            public int NoSuchColumn = 10;

            public string MappedValue;
        }
#pragma warning restore CS0649

        [Fact]
        public void ReadRows_NotIgnoredRecursiveProperty_ThrowsExcelMappingException()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NotIgnoredRecursivePropertyClass>());
        }

        [Fact]
        public void ReadRows_NotIgnoredRecursiveField_ThrowsExcelMappingException()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            Assert.Throws<ExcelMappingException>(() => sheet.ReadRow<NotIgnoredRecursiveFieldClass>());
        }

        [Fact]
        public void ReadRows_IgnoredRecursiveProperty_DoesNotDeserialize()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            IgnoredRecursivePropertyClass row1 = sheet.ReadRow<IgnoredRecursivePropertyClass>();
            Assert.Null(row1.StringValue);
            Assert.Equal("a", row1.MappedValue);
        }

        [Fact]
        public void ReadRows_IgnoredRecursiveField_DoesNotDeserialize()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");

            ExcelSheet sheet = importer.ReadSheet();
            sheet.ReadHeading();

            IgnoredRecursiveFieldClass row1 = sheet.ReadRow<IgnoredRecursiveFieldClass>();
            Assert.Null(row1.StringValue);
            Assert.Equal("a", row1.MappedValue);
        }

        private class NotIgnoredRecursivePropertyClass
        {
            public NotIgnoredRecursivePropertyClass StringValue { get; set; }

            public string MappedValue { get; set; }
        }

        private class IgnoredRecursivePropertyClass
        {
            [ExcelIgnore]
            public IgnoredRecursivePropertyClass StringValue { get; set; }

            public string MappedValue { get; set; }
        }

#pragma warning disable CS0649
        private class NotIgnoredRecursiveFieldClass
        {
            public NotIgnoredRecursiveFieldClass StringValue;

            public string MappedValue;
        }

        private class IgnoredRecursiveFieldClass
        {
            [ExcelIgnore]
            public IgnoredRecursiveFieldClass StringValue;

            public string MappedValue;
        }
#pragma warning restore CS0649
    }
}

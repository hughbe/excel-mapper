using System;
using Xunit;

namespace ExcelMapper.Tests
{
    public class ExcelImporterConfigurationTests
    {
        [Fact]
        public void RegisterMapping_ValidMappingType_Success()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                importer.Configuration.RegisterMapping<TestMap>();

                Assert.True(importer.Configuration.TryGetMapping<int>(out ExcelClassMap mapping));
                Assert.IsType<TestMap>(mapping);

                Assert.True(importer.Configuration.TryGetMapping(typeof(int), out mapping));
                Assert.IsType<TestMap>(mapping);

                Assert.IsType<TestMap>(importer.Configuration.GetMapping<int>());
                Assert.IsType<TestMap>(importer.Configuration.GetMapping(typeof(int)));
            }
        }

        [Fact]
        public void RegisterMapping_ValidMappingObject_Success()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                var map = new TestMap();
                importer.Configuration.RegisterMapping(map);

                Assert.True(importer.Configuration.TryGetMapping<int>(out ExcelClassMap mapping));
                Assert.Same(map, mapping);

                Assert.True(importer.Configuration.TryGetMapping(typeof(int), out mapping));
                Assert.Same(map, mapping);

                Assert.Same(map, importer.Configuration.GetMapping<int>());
                Assert.Same(map, importer.Configuration.GetMapping(typeof(int)));
            }
        }

        [Fact]
        public void GetMapping_NoSuchMapping_ThrowsExcelMappingException()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                Assert.False(importer.Configuration.TryGetMapping<TestMap>(out ExcelClassMap mapping));
                Assert.Null(mapping);

                Assert.False(importer.Configuration.TryGetMapping(typeof(TestMap), out mapping));
                Assert.Null(mapping);

                Assert.Throws<ExcelMappingException>(() => importer.Configuration.GetMapping<TestMap>());
                Assert.Throws<ExcelMappingException>(() => importer.Configuration.GetMapping(typeof(TestMap)));
            }
        }

        [Fact]
        public void RegisterMapping_NullMapping_ThrowsArgumentNullException()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                Assert.Throws<ArgumentNullException>("mapping", () => importer.Configuration.RegisterMapping(null));
            }
        }

        [Fact]
        public void RegisterMapping_MappingAlreadyRegistered_ThrowsExcelMappingException()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                importer.Configuration.RegisterMapping<TestMap>();
                Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterMapping<TestMap>());
                Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterMapping(new TestMap()));
            }
        }

        private class TestMap : ExcelClassMap<int> { }
    }
}

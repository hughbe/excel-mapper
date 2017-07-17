using System;
using Xunit;

namespace ExcelMapper.Tests
{
    public class ExcelImporterConfigurationTests
    {
        [Fact]
        public void RegisterClassMap_ValidClassMapType_Success()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                importer.Configuration.RegisterClassMap<TestMap>();

                Assert.True(importer.Configuration.TryGetClassMap<int>(out ExcelClassMap classMap));
                Assert.IsType<TestMap>(classMap);

                Assert.True(importer.Configuration.TryGetClassMap(typeof(int), out classMap));
                Assert.IsType<TestMap>(classMap);
            }
        }

        [Fact]
        public void RegisterClassMap_ValidClassMapObject_Success()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                var map = new TestMap();
                importer.Configuration.RegisterClassMap(map);

                Assert.True(importer.Configuration.TryGetClassMap<int>(out ExcelClassMap classMap));
                Assert.Same(map, classMap);

                Assert.True(importer.Configuration.TryGetClassMap(typeof(int), out classMap));
                Assert.Same(map, classMap);
            }
        }

        [Fact]
        public void TryGetClassMap_NullClassType_ThrowsArgumentNullException()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                ExcelClassMap classMap = null;                
                Assert.Throws<ArgumentNullException>("classType", () => importer.Configuration.TryGetClassMap(null, out classMap));
                Assert.Null(classMap);
            }
        }

        [Fact]
        public void TryGetClassMap_NoSuchClassType_ReturnsFalse()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                importer.Configuration.RegisterClassMap<OtherTestMap>();

                Assert.False(importer.Configuration.TryGetClassMap<TestMap>(out ExcelClassMap classMap));
                Assert.Null(classMap);

                Assert.False(importer.Configuration.TryGetClassMap(typeof(TestMap), out classMap));
                Assert.Null(classMap);
            }
        }

        [Fact]
        public void TryGetClassMap_NoRegisteredClassMaps_ReturnsFalse()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                Assert.False(importer.Configuration.TryGetClassMap<TestMap>(out ExcelClassMap classMap));
                Assert.Null(classMap);

                Assert.False(importer.Configuration.TryGetClassMap(typeof(TestMap), out classMap));
                Assert.Null(classMap);
            }
        }

        [Fact]
        public void RegisterClassMap_NullClassMap_ThrowsArgumentNullException()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                Assert.Throws<ArgumentNullException>("classMap", () => importer.Configuration.RegisterClassMap(null));
            }
        }

        [Fact]
        public void RegisterClassMap_ClassTypeAlreadyRegistered_ThrowsExcelMappingException()
        {
            using (var importer = Helpers.GetImporter("Primitives.xlsx"))
            {
                importer.Configuration.RegisterClassMap<TestMap>();
                Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap<TestMap>());
                Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap(new TestMap()));
            }
        }

        private class TestMap : ExcelClassMap<int> { }
        private class OtherTestMap : ExcelClassMap<int> { }
    }
}

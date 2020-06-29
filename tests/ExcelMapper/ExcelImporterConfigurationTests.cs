using System;
using System.Reflection;
using ExcelDataReader;
using Xunit;

namespace ExcelMapper.Tests
{
    public class ExcelImporterConfigurationTests
    {
        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public void SkipBlankLines_Set_GetReturnsExpected(bool value)
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            importer.Configuration.SkipBlankLines = value;
            Assert.Equal(value, importer.Configuration.SkipBlankLines);

            // Set same.
            importer.Configuration.SkipBlankLines = value;
            Assert.Equal(value, importer.Configuration.SkipBlankLines);

            // Set different.
            importer.Configuration.SkipBlankLines = !value;
            Assert.Equal(!value, importer.Configuration.SkipBlankLines);
        }

        [Fact]
        public void RegisterClassMap_InvokeDefault_Success()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            importer.Configuration.RegisterClassMap<TestMap>();

            Assert.True(importer.Configuration.TryGetClassMap<int>(out IMap classMap));
            TestMap map = Assert.IsType<TestMap>(classMap);
            Assert.Equal(FallbackStrategy.ThrowIfPrimitive, map.EmptyValueStrategy);
            Assert.Equal(typeof(int), map.Type);
            Assert.Empty(map.Properties);

            Assert.True(importer.Configuration.TryGetClassMap(typeof(int), out classMap));
            map = Assert.IsType<TestMap>(classMap);
            Assert.Equal(FallbackStrategy.ThrowIfPrimitive, map.EmptyValueStrategy);
            Assert.Equal(typeof(int), map.Type);
            Assert.Empty(map.Properties);
        }

        [Fact]
        public void RegisterClassMap_InvokeExcelClassMap_Success()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            var map = new TestMap();
            importer.Configuration.RegisterClassMap(map);

            Assert.True(importer.Configuration.TryGetClassMap<int>(out IMap classMap));
            Assert.Same(map, classMap);

            Assert.True(importer.Configuration.TryGetClassMap(typeof(int), out classMap));
            Assert.Same(map, classMap);
        }

        [Fact]
        public void RegisterClassMap_InvokeTypeIMap_Success()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            var map = new CustomIMap();
            importer.Configuration.RegisterClassMap(typeof(int), map);

            Assert.True(importer.Configuration.TryGetClassMap<int>(out IMap classMap));
            Assert.Same(map, classMap);

            Assert.True(importer.Configuration.TryGetClassMap(typeof(int), out classMap));
            Assert.Same(map, classMap);
        }

        [Fact]
        public void RegisterClassMap_NullClassType_ThrowsArgumentNullException()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            var map = new CustomIMap();
            Assert.Throws<ArgumentNullException>("classType", () => importer.Configuration.RegisterClassMap(null, map));
        }

        [Fact]
        public void RegisterClassMap_NullClassMap_ThrowsArgumentNullException()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            Assert.Throws<ArgumentNullException>("classMap", () => importer.Configuration.RegisterClassMap(null));
            Assert.Throws<ArgumentNullException>("classMap", () => importer.Configuration.RegisterClassMap(typeof(int), null));
        }

        [Fact]
        public void RegisterClassMap_ClassTypeAlreadyRegistered_ThrowsExcelMappingException()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            importer.Configuration.RegisterClassMap<TestMap>();
            Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap<TestMap>());
            Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap(new TestMap()));
            Assert.Throws<ExcelMappingException>(() => importer.Configuration.RegisterClassMap(typeof(int), new TestMap()));
        }

        [Fact]
        public void TryGetClassMap_NullClassType_ThrowsArgumentNullException()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            IMap classMap = null;
            Assert.Throws<ArgumentNullException>("classType", () => importer.Configuration.TryGetClassMap(null, out classMap));
            Assert.Null(classMap);
        }

        [Fact]
        public void TryGetClassMap_NoSuchClassType_ReturnsFalse()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            importer.Configuration.RegisterClassMap<OtherTestMap>();

            Assert.False(importer.Configuration.TryGetClassMap<TestMap>(out IMap classMap));
            Assert.Null(classMap);

            Assert.False(importer.Configuration.TryGetClassMap(typeof(TestMap), out classMap));
            Assert.Null(classMap);
        }

        [Fact]
        public void TryGetClassMap_NoRegisteredClassMaps_ReturnsFalse()
        {
            using var importer = Helpers.GetImporter("Primitives.xlsx");
            Assert.False(importer.Configuration.TryGetClassMap<TestMap>(out IMap classMap));
            Assert.Null(classMap);

            Assert.False(importer.Configuration.TryGetClassMap(typeof(TestMap), out classMap));
            Assert.Null(classMap);
        }

        private class CustomIMap : IMap
        {
            public bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo member, out object value)
            {
                throw new NotImplementedException();
            }
        }

        private class TestMap : ExcelClassMap<int>
        {
        }

        private class OtherTestMap : ExcelClassMap<int>
        {
        }
    }
}

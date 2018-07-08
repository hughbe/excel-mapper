using System;
using System.Collections.Generic;
using System.Reflection;
using ExcelMapper.Tests;
using Xunit;

namespace ExcelMapper.Utilities.Tests
{
    public class ExcelImporterUtilsTests
    {
#if !NETCOREAPP1_1
        [Fact]
        public void RegisterClassMapsInNamespace_NoAssemblyAndValidNamespaceString_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("Strings.xlsx"))
            {
                IEnumerable<ExcelClassMap> classMaps = ExcelImporterUtils.RegisterClassMapsInNamespace(importer, "ExcelMapper.Utilities.Tests");
                ExcelClassMap classMap = Assert.Single(classMaps);
                Assert.IsType<TestClassMap>(classMap);
            }
        }
#endif

        [Fact]
        public void RegisterClassMapsInNamespace_AssemblyAndValidNamespaceString_ReturnsExpected()
        {
            using (var importer = Helpers.GetImporter("Strings.xlsx"))
            {
                Assembly assembly = typeof(ExcelImporterUtilsTests).GetTypeInfo().Assembly;
                IEnumerable<ExcelClassMap> classMaps = ExcelImporterUtils.RegisterClassMapsInNamespace(importer, assembly, "ExcelMapper.Utilities.Tests");
                ExcelClassMap classMap = Assert.Single(classMaps);
                Assert.IsType<TestClassMap>(classMap);
            }
        }

        [Fact]
        public void RegisterClassMapsInNamespace_NullAssembly_ThrowsArgumentNullException()
        {
            using (var importer = Helpers.GetImporter("Strings.xlsx"))
            {
                Assert.Throws<ArgumentNullException>("assembly", () => ExcelImporterUtils.RegisterClassMapsInNamespace(importer, null, "ExcelMapper.Utilities.Tests"));
            }
        }

        [Fact]
        public void RegisterClassMapsInNamespace_NullNamespaceString_ThrowsArgumentNullException()
        {
            using (var importer = Helpers.GetImporter("Strings.xlsx"))
            {
                Assembly assembly = typeof(ExcelImporterUtils).GetTypeInfo().Assembly;
                Assert.Throws<ArgumentNullException>("namespaceString", () => ExcelImporterUtils.RegisterClassMapsInNamespace(importer, assembly, null));
            }
        }

        [Fact]
        public void RegisterClassMapsInNamespace_EmptyNamespaceString_ThrowsArgumentException()
        {
            using (var importer = Helpers.GetImporter("Strings.xlsx"))
            {
                Assembly assembly = typeof(ExcelImporterUtils).GetTypeInfo().Assembly;
                Assert.Throws<ArgumentException>("namespaceString", () => ExcelImporterUtils.RegisterClassMapsInNamespace(importer, assembly, ""));
            }
        }

        [Fact]
        public void RegisterClassMapsInNamespace_InvalidNamespaceString_ThrowsArgumentException()
        {
            using (var importer = Helpers.GetImporter("Strings.xlsx"))
            {
                Assembly assembly = typeof(ExcelImporterUtils).GetTypeInfo().Assembly;
                Assert.Throws<ArgumentException>("namespaceString", () => ExcelImporterUtils.RegisterClassMapsInNamespace(importer, assembly, "INVALID_NAMESPACE"));
            }
        }
    }

    public class TestClassMap : ExcelClassMap<TestClass>
    {
    }

    public class TestClass
    {
        public string Value { get; set; }
    }
}

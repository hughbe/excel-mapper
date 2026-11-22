using System.Reflection;
using ExcelMapper.Tests;

namespace ExcelMapper.Tests.ClassMapsNamespace;

// ExcelImporterConfiguration.RegisterClassMapsInNamespace, with a namespace of its own that holds
// one class map.

public class ExcelImporterConfigurationNamespaceTests
{
#if !NETCOREAPP1_1
    [Fact]
    public void RegisterClassMapsInNamespace_NoAssemblyAndValidNamespaceString_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        IReadOnlyList<ExcelClassMap> classMaps = importer.Configuration.RegisterClassMapsInNamespace("ExcelMapper.Tests.ClassMapsNamespace");
        ExcelClassMap classMap = Assert.Single(classMaps);
        Assert.IsType<TestClassMap>(classMap);
    }
#endif

    [Fact]
    public void RegisterClassMapsInNamespace_AssemblyAndValidNamespaceString_ReturnsExpected()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        Assembly assembly = typeof(ExcelImporterConfigurationNamespaceTests).GetTypeInfo().Assembly;
        IReadOnlyList<ExcelClassMap> classMaps = importer.Configuration.RegisterClassMapsInNamespace(assembly, "ExcelMapper.Tests.ClassMapsNamespace");
        ExcelClassMap classMap = Assert.Single(classMaps);
        Assert.IsType<TestClassMap>(classMap);
    }

    [Fact]
    public void RegisterClassMapsInNamespace_NullAssembly_ThrowsArgumentNullException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        Assert.Throws<ArgumentNullException>("assembly", () => importer.Configuration.RegisterClassMapsInNamespace(null!, "ExcelMapper.Tests.ClassMapsNamespace"));
    }

    [Fact]
    public void RegisterClassMapsInNamespace_NullNamespaceString_ThrowsArgumentNullException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        Assembly assembly = typeof(ExcelImporterConfiguration).GetTypeInfo().Assembly;
        Assert.Throws<ArgumentNullException>("namespaceString", () => importer.Configuration.RegisterClassMapsInNamespace(assembly, null!));
    }

    [Fact]
    public void RegisterClassMapsInNamespace_EmptyNamespaceString_ThrowsArgumentException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        Assembly assembly = typeof(ExcelImporterConfiguration).GetTypeInfo().Assembly;
        Assert.Throws<ArgumentException>("namespaceString", () => importer.Configuration.RegisterClassMapsInNamespace(assembly, ""));
    }

    [Fact]
    public void RegisterClassMapsInNamespace_InvalidNamespaceString_ThrowsArgumentException()
    {
        using var importer = Helpers.GetImporter("Strings.xlsx");
        Assembly assembly = typeof(ExcelImporterConfiguration).GetTypeInfo().Assembly;
        Assert.Throws<ArgumentException>("namespaceString", () => importer.Configuration.RegisterClassMapsInNamespace(assembly, "INVALID_NAMESPACE"));
    }
}

public class TestClassMap : ExcelClassMap<TestClass>
{
}

public class TestClass
{
    public string Value { get; set; } = default!;
}

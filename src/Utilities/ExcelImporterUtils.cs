using System.Linq;
using System.Reflection;

namespace ExcelMapper;

/// <summary>
/// Extension methods for <see cref="ExcelImporter"/> to provide additional functionality.
/// </summary>
public static class ExcelImporterUtils
{
    /// <summary>
    /// Registers all <see cref="ExcelClassMap"/> types found in the specified namespace within the calling assembly.
    /// </summary>
    /// <param name="importer">The importer to register the class maps with.</param>
    /// <param name="namespaceString">The namespace to search for <see cref="ExcelClassMap"/> types.</param>
    /// <returns>A collection of registered class maps.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="namespaceString"/> is null.</exception>
    /// <exception cref="ArgumentException">Thrown when <paramref name="namespaceString"/> is empty or no class maps are found in the namespace.</exception>
    public static IEnumerable<ExcelClassMap> RegisterClassMapsInNamespace(this ExcelImporter importer, string namespaceString)
    {
        return RegisterClassMapsInNamespace(importer, Assembly.GetCallingAssembly(), namespaceString);
    }

    /// <summary>
    /// Registers all <see cref="ExcelClassMap"/> types found in the specified namespace within the specified assembly.
    /// </summary>
    /// <param name="importer">The importer to register the class maps with.</param>
    /// <param name="assembly">The assembly to search for <see cref="ExcelClassMap"/> types.</param>
    /// <param name="namespaceString">The namespace to search for <see cref="ExcelClassMap"/> types.</param>
    /// <returns>A collection of registered class maps.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="assembly"/> or <paramref name="namespaceString"/> is null.</exception>
    /// <exception cref="ArgumentException">Thrown when <paramref name="namespaceString"/> is empty or no class maps are found in the namespace.</exception>
    public static IEnumerable<ExcelClassMap> RegisterClassMapsInNamespace(this ExcelImporter importer, Assembly assembly, string namespaceString)
    {
        ArgumentNullException.ThrowIfNull(assembly);

        ArgumentNullException.ThrowIfNull(namespaceString);
        ArgumentException.ThrowIfNullOrEmpty(namespaceString);

        var classMapTypes = assembly
            .GetTypes()
            .Where(type => typeof(ExcelClassMap).IsAssignableFrom(type) && type.Namespace == namespaceString);

        var classMaps = new List<ExcelClassMap>();
        foreach (var type in classMapTypes)
        {
            if (Activator.CreateInstance(type) is ExcelClassMap classMap)
            {
                classMaps.Add(classMap);
                importer.Configuration.RegisterClassMap(classMap);
            }
        }

        if (classMaps.Count == 0)
        {
            throw new ArgumentException($"No ExcelClassMap types found in the namespace \"{namespaceString}\" in the assembly \"{assembly}\".", nameof(namespaceString));
        }

        return classMaps;
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ExcelMapper;

public static class ExcelImporterUtils
{
    public static IEnumerable<ExcelClassMap> RegisterClassMapsInNamespace(this ExcelImporter importer, string namespaceString)
    {
        return RegisterClassMapsInNamespace(importer, Assembly.GetCallingAssembly(), namespaceString);
    }

    public static IEnumerable<ExcelClassMap> RegisterClassMapsInNamespace(this ExcelImporter importer, Assembly assembly, string namespaceString)
    {
        ArgumentNullException.ThrowIfNull(assembly);

        ArgumentNullException.ThrowIfNull(namespaceString);
        ArgumentException.ThrowIfNullOrEmpty(namespaceString);

        var classMapTypes = assembly
            .GetTypes()
            .Where(type => typeof(ExcelClassMap).IsAssignableFrom(type) && type.Namespace == namespaceString);

        var classMaps = classMapTypes.Select(Activator.CreateInstance).OfType<ExcelClassMap>().ToArray();
        if (classMaps.Length == 0)
        {
            throw new ArgumentException($"No ExcelClassMap types found in the namespace \"{namespaceString}\" in the assembly \"{assembly}\".", nameof(namespaceString));
        }

        foreach (ExcelClassMap classMap in classMaps)
        {
            importer.Configuration.RegisterClassMap(classMap);
        }

        return classMaps;
    }
}

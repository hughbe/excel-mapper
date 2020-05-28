using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ExcelMapper
{
    public static class ExcelImporterUtils
    {
        public static IEnumerable<ExcelClassMap> RegisterClassMapsInNamespace(this ExcelImporter importer, string namespaceString)
        {
            return RegisterClassMapsInNamespace(importer, Assembly.GetCallingAssembly(), namespaceString);
        }

        public static IEnumerable<ExcelClassMap> RegisterClassMapsInNamespace(this ExcelImporter importer, Assembly assembly, string namespaceString)
        {
            if (assembly == null)
            {
                throw new ArgumentNullException(nameof(assembly));
            }

            if (namespaceString == null)
            {
                throw new ArgumentNullException(nameof(namespaceString));
            }

            if (namespaceString.Length == 0)
            {
                throw new ArgumentException("The namespace cannot be empty.", nameof(namespaceString));
            }

            IEnumerable<Type> classMapTypes = assembly
                .GetTypes()
                .Where(type => typeof(ExcelClassMap).IsAssignableFrom(type) && type.Namespace == namespaceString);

            ExcelClassMap[] classMaps = classMapTypes.Select(Activator.CreateInstance).OfType<ExcelClassMap>().ToArray();
            if (classMaps.Length == 0)
            {
                throw new ArgumentException($"No classmaps found in the namespace \"{namespaceString}\" in the assembly \"{assembly}\".", nameof(namespaceString));
            }

            foreach (ExcelClassMap classMap in classMaps)
            {
                importer.Configuration.RegisterClassMap(classMap);
            }

            return classMaps;
        }
    }
}

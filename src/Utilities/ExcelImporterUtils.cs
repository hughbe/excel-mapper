using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ExcelMapper
{
    public static class ExcelImporterUtils
    {
#if NETSTANDARD2_0
        public static IEnumerable<ExcelClassMap> RegisterClassMapsInNamespace(this ExcelImporter importer,
            string namespaceString)
        {
            return RegisterClassMapsInNamespace(importer, Assembly.GetExecutingAssembly(), namespaceString);
        }
#endif

        public static IEnumerable<ExcelClassMap> RegisterClassMapsInNamespace(this ExcelImporter importer,
            Assembly assembly,
            string namespaceString)
        {
            if (namespaceString == null)
            {
                throw new ArgumentNullException(nameof(namespaceString));
            }

            if (namespaceString.Length == 0)
            {
                throw new ArgumentException("The namespace cannot be empty.", nameof(namespaceString));
            }

            var classes = assembly
                .GetTypes()
                .Where(t => t.IsAssignableFrom(typeof(ExcelClassMap)) && t.Namespace == namespaceString);

            var objects = classes.Select(Activator.CreateInstance).OfType<ExcelClassMap>().ToList();

            foreach (var o in objects)
                importer.Configuration.RegisterClassMap(o);

            return objects;
        }
    }
}
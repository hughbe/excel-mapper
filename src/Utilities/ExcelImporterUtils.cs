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
            string nameSpace)
        {
            return RegisterClassMapsInNamespace(importer, Assembly.GetExecutingAssembly(), nameSpace);
        }
#endif

        public static IEnumerable<ExcelClassMap> RegisterClassMapsInNamespace(this ExcelImporter importer,
            Assembly assembly,
            string nameSpace)
        {
            if (nameSpace == null)
            {
                throw new ArgumentNullException(nameof(nameSpace));
            }

            if (nameSpace.Length == 0)
            {
                throw new ArgumentException("The namespace cannot be empty.", nameof(nameSpace));
            }

            var classes = assembly
                .GetTypes()
                .Where(t => t.IsAssignableFrom(typeof(ExcelClassMap)) && t.Namespace == nameSpace);

            var objects = classes.Select(Activator.CreateInstance).OfType<ExcelClassMap>().ToList();

            foreach (var o in objects)
                importer.Configuration.RegisterClassMap(o);

            return objects;
        }
    }
}
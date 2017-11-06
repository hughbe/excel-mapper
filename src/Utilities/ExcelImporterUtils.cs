using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ExcelMapper
{
    public static class ExcelImporterUtils
    {
#if NETSTANDARD2_0
        public static List<ExcelClassMap> RegisterMapperClassesByNamespace(this ExcelImporter importer,
            string namespacestr)
        {
            var classes = Assembly.GetExecutingAssembly()
                .GetTypes().AsParallel()
                .Where(t => t.IsClass && t.Namespace == namespacestr);

            var objects = classes.AsParallel().Select(Activator.CreateInstance)
                .AsParallel().OfType<ExcelClassMap>().ToList();

            foreach (var o in objects)
                importer.Configuration.RegisterClassMap(o);

            return objects;
        }
#endif
    }
}
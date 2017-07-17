using System;
using System.Collections.Generic;

namespace ExcelMapper
{
    public class ExcelImporterConfiguration
    {
        private List<ExcelClassMap> ClassMaps { get; } = new List<ExcelClassMap>();

        internal ExcelImporterConfiguration() { }

        public bool TryGetClassMap(Type classType, out ExcelClassMap classMap)
        {
            if (classType == null)
            {
                throw new ArgumentNullException(nameof(classType));
            }

            foreach (ExcelClassMap registeredMap in ClassMaps)
            {
                if (registeredMap.Type == classType)
                {
                    classMap = registeredMap;
                    return true;
                }
            }

            classMap = null;
            return false;
        }

        public bool TryGetClassMap<T>(out ExcelClassMap mapping) => TryGetClassMap(typeof(T), out mapping);

        public void RegisterClassMap<T>() where T : ExcelClassMap, new()
        {
            T classMap = Activator.CreateInstance<T>();
            RegisterClassMap(classMap);
        }

        public void RegisterClassMap(ExcelClassMap classMap)
        {
            if (classMap == null)
            {
                throw new ArgumentNullException(nameof(classMap));
            }

            if (TryGetClassMap(classMap.Type, out ExcelClassMap registeredClassMap))
            {
                throw new ExcelMappingException($"Class map already type \"{classMap.Type.FullName}\"");
            }

            ClassMaps.Add(classMap);
        }

        public Func<ExcelSheet, bool> HasHeading { get; set; }
    }
}

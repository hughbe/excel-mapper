using System;
using System.Collections.Generic;

namespace ExcelMapper
{
    public class ExcelImporterConfiguration
    {
        private List<ExcelClassMap> Mappings { get; } = new List<ExcelClassMap>();

        internal ExcelImporterConfiguration() { }

        public bool TryGetMapping(Type type, out ExcelClassMap mapping)
        {
            if (type == null)
            {
                throw new ArgumentNullException(nameof(type));
            }

            foreach (ExcelClassMap registeredMapping in Mappings)
            {
                if (registeredMapping.Type == type)
                {
                    mapping = registeredMapping;
                    return true;
                }
            }

            mapping = null;
            return false;
        }

        public bool TryGetMapping<T>(out ExcelClassMap mapping) => TryGetMapping(typeof(T), out mapping);

        public ExcelClassMap GetMapping(Type type)
        {
            if (type == null)
            {
                throw new ArgumentNullException(nameof(type));
            }

            if (!TryGetMapping(type, out ExcelClassMap mapping))
            {
                throw new ExcelMappingException($"No registered mapping for type \"{type.FullName}\".");
            }

            return mapping;
        }

        public ExcelClassMap GetMapping<T>() => GetMapping(typeof(T));

        public void RegisterMapping<T>() where T : ExcelClassMap, new()
        {
            T mapping = Activator.CreateInstance<T>();
            RegisterMapping(mapping);
        }

        public void RegisterMapping(ExcelClassMap mapping)
        {
            if (mapping == null)
            {
                throw new ArgumentNullException(nameof(mapping));
            }

            if (TryGetMapping(mapping.Type, out ExcelClassMap registeredMapping))
            {
                throw new ExcelMappingException($"Mapping already type \"{mapping.Type.FullName}\"");
            }

            Mappings.Add(mapping);
        }

        public Func<ExcelSheet, bool> HasHeading { get; set; }
    }
}

using System;
using System.Collections.Generic;

namespace ExcelMapper
{
    /// <summary>
    /// Configuration options for importing Excel documents. The main function is to handle registering
    /// known class maps.
    /// </summary>
    public class ExcelImporterConfiguration
    {
        private List<ExcelClassMap> ClassMaps { get; } = new List<ExcelClassMap>();

        internal ExcelImporterConfiguration() { }

        /// <summary>
        /// Gets the class map registered for an object of the given type.
        /// </summary>
        /// <param name="classType">The type of the object to get the class map for.</param>
        /// <param name="classMap">The class map for the given type if it exists, else null.</param>
        /// <returns>True if a class map exists for the given type, else false.</returns>
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

        /// <summary>
        /// Gets the class map registered for an object of the given type.
        /// </summary>
        /// <typeparam name="T">The type of the object to get the class map for.</typeparam>
        /// <param name="classMap">The class map for the given type if it exists, else null.</param>
        /// <returns>True if a class map exists for the given type, else false.</returns>
        public bool TryGetClassMap<T>(out ExcelClassMap classMap) => TryGetClassMap(typeof(T), out classMap);

        /// <summary>
        /// Registers a class map of the given type to be used when mapping a row to an object.
        /// </summary>
        /// <typeparam name="T">The type of the class map to use.</typeparam>
        public void RegisterClassMap<T>() where T : ExcelClassMap, new()
        {
            T classMap = Activator.CreateInstance<T>();
            RegisterClassMap(classMap);
        }

        /// <summary>
        /// Registers the given class map to be used when mapping a row to an object.
        /// </summary>
        /// <param name="classMap">The class map to use.</param>
        public void RegisterClassMap(ExcelClassMap classMap)
        {
            if (classMap == null)
            {
                throw new ArgumentNullException(nameof(classMap));
            }

            if (TryGetClassMap(classMap.Type, out ExcelClassMap _))
            {
                throw new ExcelMappingException($"Class map already type \"{classMap.Type.FullName}\"");
            }

            ClassMaps.Add(classMap);
        }
    }
}

using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using ExcelMapper.Abstractions;

namespace ExcelMapper;

/// <summary>
/// Configuration options for importing Excel documents. The main function is to handle registering
/// known class maps.
/// </summary>
public class ExcelImporterConfiguration
{
    private Dictionary<Type, IMap> ClassMaps { get; } = [];


    /// <summary>
    ///  Gets or sets whether blank lines should be skipped during reading.
    ///  This may have performance implications so is off by default.
    /// </summary>
    public bool SkipBlankLines { get; set; }


    internal ExcelImporterConfiguration() { }

    /// <summary>
    /// Gets the class map registered for an object of the given type.
    /// </summary>
    /// <param name="classType">The type of the object to get the class map for.</param>
    /// <param name="classMap">The class map for the given type if it exists, else null.</param>
    /// <returns>True if a class map exists for the given type, else false.</returns>
    public bool TryGetClassMap(Type classType, [NotNullWhen(true)] out IMap? classMap)
    {
        if (classType == null)
        {
            throw new ArgumentNullException(nameof(classType));
        }

        return ClassMaps.TryGetValue(classType, out classMap);
    }

    /// <summary>
    /// Gets the class map registered for an object of the given type.
    /// </summary>
    /// <typeparam name="T">The type of the object to get the class map for.</typeparam>
    /// <param name="classMap">The class map for the given type if it exists, else null.</param>
    /// <returns>True if a class map exists for the given type, else false.</returns>
    public bool TryGetClassMap<T>([NotNullWhen(true)] out IMap? classMap) => TryGetClassMap(typeof(T), out classMap);

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

        RegisterClassMap(classMap.Type, classMap);
    }

    private void ValidateMap(IMap map)
    {
        if (map is ExcelClassMap excelClassMap)
        {
            foreach (var propertyMap in excelClassMap.Properties)
            {
                ValidateMap(propertyMap.Map);
            }
        }
        else if (map is IEnumerableIndexerMap enumerableIndexerMap)
        {
            foreach (var itemMap in enumerableIndexerMap.Values)
            {
                ValidateMap(itemMap.Value);
            }
        }
        else if (map is IMultidimensionalIndexerMap multidimensionalIndexerMap)
        {
            foreach (var itemMap in multidimensionalIndexerMap.Values)
            {
                ValidateMap(itemMap.Value);
            }
        }
        else if (map is IDictionaryIndexerMap dictionaryIndexerMap)
        {
            foreach (var itemMap in dictionaryIndexerMap.Values)
            {
                ValidateMap(itemMap.Value);
            }
        }
        else if (map is IValuePipeline pipeline)
        {
            if (pipeline.CellValueMappers.Count == 0)
            {
                throw new ExcelMappingException("Cannot register a class map with a property that has no CellValueMappers defined.");
            }
        }
    }

    /// <summary>
    /// Registers the given class map to be used when mapping a row to an object.
    /// </summary>
    /// <param name="classType">The type of the class.</param>
    /// <param name="classMap">The class map to use.</param>
    public void RegisterClassMap(Type classType, IMap classMap)
    {
        if (classType == null)
        {
            throw new ArgumentNullException(nameof(classType));
        }
        if (classMap == null)
        {
            throw new ArgumentNullException(nameof(classMap));
        }
        if (ClassMaps.ContainsKey(classType))
        {
            throw new ExcelMappingException($"Class map already exists for type \"{classType.FullName}\"");
        }

        ValidateMap(classMap);
        ClassMaps.Add(classType, classMap);
    }
}

using System.Collections.Concurrent;
using System.Diagnostics.CodeAnalysis;

namespace ExcelMapper;

/// <summary>
/// Configuration options for importing Excel documents. The main function is to handle registering
/// known class maps.
/// </summary>
public class ExcelImporterConfiguration
{
    private ConcurrentDictionary<Type, IMap> ClassMaps { get; } = new();

    private int _maxColumnsPerSheet = 10000;

    /// <summary>
    /// Gets or sets the maximum number of columns allowed per sheet. Default is 10000.
    /// This limit prevents denial of service attacks from malicious Excel files with excessive columns.
    /// Set to int.MaxValue to disable the limit (not recommended for untrusted files).
    /// Excel .xlsx files support up to 16,384 columns (XFD).
    /// </summary>
    public int MaxColumnsPerSheet
    {
        get => _maxColumnsPerSheet;
        set
        {
            ArgumentOutOfRangeException.ThrowIfNegativeOrZero(value);
            _maxColumnsPerSheet = value;
        }
    }

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
        ArgumentNullException.ThrowIfNull(classType);

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
    /// Registers a class map of the given type to be used when mapping a row to an
    /// object.
    /// </summary> <typeparam name="T">The type of the class to map.</typeparam>
    /// <param name="classMapFactory">A factory that configures and returns the class map to use.</param>
    public void RegisterClassMap<T>(Action<ExcelClassMap<T>> classMapFactory)
    {
        ArgumentNullException.ThrowIfNull(classMapFactory);

        var classMapInstance = new ExcelClassMap<T>();
        classMapFactory(classMapInstance);
        RegisterClassMap(typeof(T), classMapInstance);
    }

    /// <summary>
    /// Registers the given class map to be used when mapping a row to an object.
    /// </summary>
    /// <param name="classMap">The class map to use.</param>
    public void RegisterClassMap(ExcelClassMap classMap)
    {
        ArgumentNullException.ThrowIfNull(classMap);

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
            if (pipeline.Mappers.Count == 0)
            {
                throw new ExcelMappingException("Cannot register a class map with a property that has no Mappers defined.");
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
        ArgumentNullException.ThrowIfNull(classType);
        ArgumentNullException.ThrowIfNull(classMap);

        ValidateMap(classMap);
        
        if (!ClassMaps.TryAdd(classType, classMap))
        {
            throw new ExcelMappingException($"Class map already exists for type \"{classType.FullName}\"");
        }
    }
}

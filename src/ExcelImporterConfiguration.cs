using System.Collections.Concurrent;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;

namespace ExcelMapper;

/// <summary>
/// Configuration options for importing Excel documents. The main function is to handle registering
/// known class maps.
/// </summary>
public class ExcelImporterConfiguration
{
    private ConcurrentDictionary<Type, ExcelClassMap> ClassMaps { get; } = new();

    // Types whose maps were created by auto-mapping rather than registered by the user. These can be
    // replaced by a map the user registers later.
    private ConcurrentDictionary<Type, ExcelClassMap> AutoMappedClassMaps { get; } = new();

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
            ThrowHelpers.ThrowIfNegativeOrZero(value);
            _maxColumnsPerSheet = value;
        }
    }

    /// <summary>
    ///  Gets or sets whether blank lines should be skipped during reading.
    ///  This may have performance implications so is off by default.
    /// </summary>
    public bool SkipBlankLines { get; set; }

    /// <summary>
    /// Gets or sets whether data annotations should be validated during mapping.
    /// Default is false.
    /// </summary>
    public bool ValidateDataAnnotations { get; set; }

    /// <summary>
    /// Gets or sets whether spaces, underscores, hyphens and dots are ignored when a column can't be
    /// found by its exact name. For example, a property named <c>FirstName</c> then matches a column
    /// named "First Name", "first_name" or "First-Name". Exact matches are always preferred.
    /// Default is false.
    /// </summary>
    public bool IgnoreSeparatorsInColumnNames { get; set; }

    /// <summary>
    /// Gets or sets the culture used to map cells: to parse text into numbers, dates and times,
    /// and to convert numbers and dates to text. When null (the default), the current culture of
    /// the thread reading the rows is used. Set this (for example to
    /// <see cref="CultureInfo.InvariantCulture"/>) so results don't depend on the machine's
    /// settings. A format provider set on an individual member takes precedence. The culture is
    /// given to the built-in conversions, not set as the thread's culture: converters get it from
    /// <see cref="CellContext.FormatProvider"/>.
    /// </summary>
    public CultureInfo? Culture { get; set; }


    internal ExcelImporterConfiguration() { }

    /// <summary>
    /// Gets the class map registered for an object of the given type.
    /// </summary>
    /// <param name="classType">The type of the object to get the class map for.</param>
    /// <param name="classMap">The class map for the given type if it exists, else null.</param>
    /// <returns>True if a class map exists for the given type, else false.</returns>
    public bool TryGetClassMap(Type classType, [NotNullWhen(true)] out ExcelClassMap? classMap)
    {
        ThrowHelpers.ThrowIfNull(classType, nameof(classType));

        return ClassMaps.TryGetValue(classType, out classMap);
    }

    /// <summary>
    /// Gets the class map registered for an object of the given type.
    /// </summary>
    /// <typeparam name="T">The type of the object to get the class map for.</typeparam>
    /// <param name="classMap">The class map for the given type if it exists, else null.</param>
    /// <returns>True if a class map exists for the given type, else false.</returns>
    public bool TryGetClassMap<T>([NotNullWhen(true)] out ExcelClassMap? classMap) => TryGetClassMap(typeof(T), out classMap);

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
    public void RegisterClassMap<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(Action<ExcelClassMap<T>> classMapFactory)
    {
        ThrowHelpers.ThrowIfNull(classMapFactory, nameof(classMapFactory));

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
        ThrowHelpers.ThrowIfNull(classMap, nameof(classMap));

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
        else if (map is ValuePipeline pipeline)
        {
            if (pipeline.MapperCount == 0)
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
    public void RegisterClassMap(Type classType, ExcelClassMap classMap)
    {
        ThrowHelpers.ThrowIfNull(classType, nameof(classType));
        ThrowHelpers.ThrowIfNull(classMap, nameof(classMap));

        ValidateMap(classMap);
        
        if (!ClassMaps.TryAdd(classType, classMap))
        {
            // A map created implicitly by auto-mapping (for example by an earlier call to ReadRows)
            // can be replaced by one the user registers.
            if (AutoMappedClassMaps.TryRemove(classType, out var autoMappedClassMap) &&
                ClassMaps.TryUpdate(classType, classMap, autoMappedClassMap))
            {
                return;
            }

            throw new ExcelMappingException($"Class map already exists for type \"{classType.FullName}\"");
        }
    }

    /// <summary>
    /// Registers a map created by auto-mapping. Unlike a map registered by the user, this can be
    /// replaced by a later call to RegisterClassMap.
    /// </summary>
    internal ExcelClassMap RegisterAutoMappedClassMap(Type classType, ExcelClassMap classMap)
    {
        var registered = ClassMaps.GetOrAdd(classType, classMap);
        if (ReferenceEquals(registered, classMap))
        {
            AutoMappedClassMaps[classType] = classMap;
        }

        return registered;
    }
}

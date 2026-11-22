using System.Collections.Concurrent;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Reflection;
using System.Runtime.CompilerServices;

namespace ExcelMapper;

/// <summary>
/// Configuration options for importing Excel documents, and the class maps registered for types.
/// </summary>
/// <remarks>
/// A sheet takes the settings (<see cref="Culture"/>, <see cref="SkipBlankLines"/>,
/// <see cref="ValidateDataAnnotations"/>, <see cref="IgnoreSeparatorsInColumnNames"/> and
/// <see cref="MaxColumnsPerSheet"/>) when it starts reading, at its heading or first row, so
/// changing them part way through doesn't change how the rest of the sheet is read. Sheets and
/// regions read afterwards use the new settings.
/// </remarks>
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
    /// and to convert numbers and dates to text. Default is <see cref="CultureInfo.InvariantCulture"/>,
    /// so results don't depend on the settings of the machine reading the file: "1.5" is one and a
    /// half everywhere. Set it to the culture the file was written in, such as
    /// <c>new CultureInfo("de-DE")</c> for "1,5", or to <see cref="CultureInfo.CurrentCulture"/>
    /// for the machine's culture. A format provider set on an individual member takes precedence.
    /// The culture is given to the built-in conversions, not set as the thread's culture:
    /// converters get it from <see cref="CellContext.FormatProvider"/>.
    /// </summary>
    /// <exception cref="ArgumentNullException">The value is null.</exception>
    public CultureInfo Culture
    {
        get => _culture;
        set
        {
            ThrowHelpers.ThrowIfNull(value);
            _culture = value;
        }
    }

    private CultureInfo _culture = CultureInfo.InvariantCulture;


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
    /// Registers a class map for <typeparamref name="T"/>, configured by <paramref name="configure"/>.
    /// </summary>
    /// <typeparam name="T">The type of the class to map.</typeparam>
    /// <param name="configure">Configures the class map, which starts with no members: call
    /// <see cref="ExcelClassMap{T}.AutoMap"/> to start from what auto-mapping would map.</param>
    public void RegisterClassMap<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(Action<ExcelClassMap<T>> configure)
    {
        ThrowHelpers.ThrowIfNull(configure, nameof(configure));

        var classMap = new ExcelClassMap<T>();
        configure(classMap);
        RegisterClassMap(classMap);
    }

    /// <summary>
    /// Registers the given class map to be used when mapping rows to its type
    /// (<see cref="ExcelClassMap.Type"/>). It replaces a class map auto-mapping registered for the
    /// type, for example in an earlier <c>ReadRows</c>.
    /// </summary>
    /// <param name="classMap">The class map to use.</param>
    /// <exception cref="ExcelMappingException">A class map is already registered for the type.</exception>
    public void RegisterClassMap(ExcelClassMap classMap)
    {
        ThrowHelpers.ThrowIfNull(classMap, nameof(classMap));
        ValidateMap(classMap);

        var classType = classMap.Type;
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
    /// Registers every <see cref="ExcelClassMap"/> type in a namespace of the calling assembly,
    /// creating each with its parameterless constructor.
    /// </summary>
    /// <param name="namespaceString">The namespace to search for <see cref="ExcelClassMap"/> types.</param>
    /// <returns>The class maps registered.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="namespaceString"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="namespaceString"/> is empty or no class maps are found in the namespace.</exception>
    [RequiresUnreferencedCode("Finds class maps by scanning the assembly's types and creates them through reflection, so the trimmer may remove them. Register class maps directly with RegisterClassMap instead.")]
    [MethodImpl(MethodImplOptions.NoInlining)] // So that the calling assembly is the caller's.
    public IReadOnlyList<ExcelClassMap> RegisterClassMapsInNamespace(string namespaceString)
        => RegisterClassMapsInNamespace(Assembly.GetCallingAssembly(), namespaceString);

    /// <summary>
    /// Registers every <see cref="ExcelClassMap"/> type in a namespace of an assembly, creating
    /// each with its parameterless constructor.
    /// </summary>
    /// <param name="assembly">The assembly to search for <see cref="ExcelClassMap"/> types.</param>
    /// <param name="namespaceString">The namespace to search for <see cref="ExcelClassMap"/> types.</param>
    /// <returns>The class maps registered.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="assembly"/> or <paramref name="namespaceString"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="namespaceString"/> is empty or no class maps are found in the namespace.</exception>
    [RequiresUnreferencedCode("Finds class maps by scanning the assembly's types and creates them through reflection, so the trimmer may remove them. Register class maps directly with RegisterClassMap instead.")]
    public IReadOnlyList<ExcelClassMap> RegisterClassMapsInNamespace(Assembly assembly, string namespaceString)
    {
        ThrowHelpers.ThrowIfNull(assembly, nameof(assembly));
        ThrowHelpers.ThrowIfNullOrEmpty(namespaceString, nameof(namespaceString));

        var classMaps = new List<ExcelClassMap>();
        foreach (var type in assembly.GetTypes())
        {
            if (typeof(ExcelClassMap).IsAssignableFrom(type) && type.Namespace == namespaceString &&
                Activator.CreateInstance(type) is ExcelClassMap classMap)
            {
                classMaps.Add(classMap);
                RegisterClassMap(classMap);
            }
        }

        if (classMaps.Count == 0)
        {
            throw new ArgumentException($"No ExcelClassMap types found in the namespace \"{namespaceString}\" in the assembly \"{assembly}\".", nameof(namespaceString));
        }

        return classMaps;
    }

    private static void ValidateMap(IMap map)
        => MapTree.VisitCellMaps(map, null, (cellMap, _) =>
        {
            if (cellMap is ValuePipeline { HasMappers: false })
            {
                throw new ExcelMappingException("Cannot register a class map with a property that has no Mappers defined.");
            }
        });

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

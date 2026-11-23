using System.Collections.Concurrent;
using System.Diagnostics.CodeAnalysis;
using ExcelMapper.AutoMapping;

namespace ExcelMapper;

/// <summary>
/// Gives class maps of the types it registers, for reading rows with native AOT, which can't create
/// the generic code auto-mapping needs for types it finds at run time. A context registers the types
/// rows are mapped to and the types of their members (see <see cref="RegisterTypes"/>), and
/// <see cref="GetClassMap{T}"/> auto-maps a type with them, as <c>ReadRows&lt;T&gt;()</c> does.
/// Read rows with the class map, for example with <see cref="ExcelSheet.ReadRows{T}(ExcelClassMap{T})"/>.
/// </summary>
/// <remarks>
/// A context is meant to be created once and shared: its class maps are created the first time
/// they're asked for, and can't be changed.
/// </remarks>
public abstract class ExcelMappingContext
{
    private readonly ConcurrentDictionary<Type, ExcelClassMap> _classMaps = new();
    private readonly object _registerLock = new();
    private HashSet<Type>? _rowTypes;
    private HashSet<(Type, Type, Type?)> _registeredTypes = [];

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelMappingContext"/> class.
    /// </summary>
    protected ExcelMappingContext()
    {
    }

    /// <summary>
    /// Registers the types rows are mapped to (<see cref="ExcelTypeRegistrar.Row{T}"/>) and every
    /// type auto-mapping meets in their members: their own types, nullable types, enums, and the
    /// element types of collections and dictionaries. Called once, the first time a class map is
    /// asked for.
    /// </summary>
    /// <param name="types">Registers the types.</param>
    protected abstract void RegisterTypes(ExcelTypeRegistrar types);

    /// <summary>
    /// Gets the class map of <typeparamref name="T"/>, auto-mapped from its members and their
    /// attributes as <c>ReadRows&lt;T&gt;()</c> does. The class map can't be changed.
    /// </summary>
    /// <typeparam name="T">A type registered with <see cref="ExcelTypeRegistrar.Row{T}"/>.</typeparam>
    /// <returns>The class map.</returns>
    /// <exception cref="InvalidOperationException"><typeparamref name="T"/> isn't registered as a row type.</exception>
    /// <exception cref="ExcelMappingException"><typeparamref name="T"/> can't be auto-mapped.</exception>
    public ExcelClassMap<T> GetClassMap<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>()
    {
        if (!GetRowTypes().Contains(typeof(T)))
        {
            throw new InvalidOperationException($"Type {TypeNames.Get(typeof(T))} isn't registered as a row type by {TypeNames.Get(GetType())}.");
        }

        if (!_classMaps.TryGetValue(typeof(T), out var classMap))
        {
            classMap = _classMaps.GetOrAdd(typeof(T), CreateClassMap<T>(GetType(), _registeredTypes));
        }

        return (ExcelClassMap<T>)classMap;
    }

    // Auto-mapping finds T's members and attributes by reflection, which native AOT supports for
    // the members ExcelClassMap<T> keeps, and creates generic types for the types it finds, which
    // the context registered (see GenericTypes).
    [UnconditionalSuppressMessage("Trimming", "IL2026", Justification = "The members of T are kept by its annotation, and the types of its members by the annotations of the ExcelTypeRegistrar methods that registered them.")]
    [UnconditionalSuppressMessage("AOT", "IL3050", Justification = "The generic types auto-mapping needs for T's members are created with the instantiations the context registered.")]
    private static ExcelClassMap<T> CreateClassMap<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(Type contextType, HashSet<(Type, Type, Type?)> registeredTypes)
    {
        // With GenericTypes.RequireRegisteredTypes on (in the tests), this throws if the context
        // doesn't register a type its class map needs (see GenericTypes.Require).
        using var scope = GenericTypes.EnterContext(contextType, registeredTypes);
        if (!AutoMapper.TryCreateClassMap<T>(EmptyValueStrategy.Error, out var classMap))
        {
            throw new ExcelMappingException($"Cannot auto-map type {TypeNames.Get(typeof(T))}.");
        }

        classMap.Freeze();
        return classMap;
    }

    private HashSet<Type> GetRowTypes()
    {
        if (_rowTypes is { } rowTypes)
        {
            return rowTypes;
        }

        lock (_registerLock)
        {
            if (_rowTypes is null)
            {
                var registrar = new ExcelTypeRegistrar();
                RegisterTypes(registrar);
                _registeredTypes = registrar.RegisteredTypes;
                _rowTypes = registrar.RowTypes;
            }

            return _rowTypes;
        }
    }
}

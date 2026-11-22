using System.Collections.Concurrent;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using System.Linq;
using System.Numerics;
using ExcelMapper.Mappers;

namespace ExcelMapper.AutoMapping;

/// <summary>
/// How members of a type are mapped, decided once per type: its built-in conversion, if it has
/// one, and the rules for its elements (<see cref="Collection"/>) or keys and values
/// (<see cref="Dictionary"/>). <see cref="AutoMapper"/> tries them in order: a value, a
/// dictionary, a collection, then an object with its own members.
/// </summary>
/// <remarks>
/// The rules are the only place where auto-mapping makes generic types at run time: each is
/// created for its type's arguments, so mapping a member calls them without reflection.
/// </remarks>
[RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
[RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
internal abstract class TypeRule
{
    private static readonly ConcurrentDictionary<Type, TypeRule> s_rules = new();

    public static TypeRule For(Type type)
        => GetOrAdd(type, static t => (TypeRule)GenericTypes.Create(typeof(TypeRule<>), t));

    // One rule for each type, however it's created.
    private protected static TypeRule GetOrAdd(Type type, Func<Type, TypeRule> create)
        => s_rules.GetOrAdd(type, create);

    /// <summary>
    /// Gets the rule for the type's elements, or null if it isn't a collection.
    /// </summary>
    public abstract CollectionRule? Collection { get; }

    /// <summary>
    /// Gets the rule for the type's keys and values, or null if it isn't a dictionary.
    /// </summary>
    public abstract DictionaryRule? Dictionary { get; }

    /// <summary>
    /// Auto-maps a member of the type.
    /// </summary>
    public abstract bool TryAutoMapMember(MemberInfo member, EmptyValueStrategy emptyValueStrategy, HashSet<Type>? typeStack, [NotNullWhen(true)] out IMap? map);

    /// <summary>
    /// Gets the class map for a nested object mapped through its members, such as
    /// <c>Map(p => p.Address.Street)</c>.
    /// </summary>
    public abstract ExcelClassMap GetOrCreateNestedMap(IMap parentMap, object? context, EmptyValueStrategy emptyValueStrategy);

}

[RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
[RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
internal sealed class TypeRule<T> : TypeRule
{
    private static TypeRule<T>? s_instance;

    /// <summary>
    /// Gets the rule for <typeparamref name="T"/>, created without reflection: native AOT has the
    /// code of a rule whose type is known when the code calling it is compiled.
    /// </summary>
    public static TypeRule<T> Instance => s_instance ??= (TypeRule<T>)GetOrAdd(typeof(T), static _ => new TypeRule<T>());

    /// <summary>
    /// Gets the adapter for converters and fallbacks of the underlying type of a nullable
    /// <typeparamref name="T"/> (see <see cref="UnderlyingTypeAdapter{T}"/>), or null if
    /// <typeparamref name="T"/> isn't nullable.
    /// </summary>
    public UnderlyingTypeAdapter<T>? NullableAdapter
    {
        get
        {
            if (!s_nullableAdapterCreated)
            {
                s_nullableAdapter = CreateNullableAdapter();
                s_nullableAdapterCreated = true;
            }

            return s_nullableAdapter;
        }
    }

    private static UnderlyingTypeAdapter<T>? CreateNullableAdapter()
        => Nullable.GetUnderlyingType(typeof(T)) is { } underlyingType
            ? (UnderlyingTypeAdapter<T>)GenericTypes.Create(typeof(NullableAdapter<>), underlyingType)
            : null;

    // Created when a converter or fallback of the underlying type needs adapting, rather than for
    // every nullable type, as it may be made with MakeGenericType.
    private static UnderlyingTypeAdapter<T>? s_nullableAdapter;
    private static bool s_nullableAdapterCreated;

    // Each is found when it's first needed: null until then, and a sentinel if there's none.
    private CollectionRule? _collection;
    private DictionaryRule? _dictionary;
    private Action<IList<ICellMapper<T>>>? _addDefaultMappers;

    private static readonly CollectionRule s_noCollection = new CollectionRule<object>();
    private static readonly DictionaryRule s_noDictionary = new DictionaryRule<string, object>();
    private static readonly Action<IList<ICellMapper<T>>> s_noDefaultMappers = static _ => { };

    public override CollectionRule? Collection
    {
        get
        {
            var rule = _collection ??= typeof(T).GetElementTypeOrEnumerableType() is { } elementType
                ? (CollectionRule)GenericTypes.Create(typeof(CollectionRule<>), elementType)
                : s_noCollection;
            return ReferenceEquals(rule, s_noCollection) ? null : rule;
        }
    }

    public override DictionaryRule? Dictionary
    {
        get
        {
            var rule = _dictionary ??= DictionaryRule.TryGetKeyValueTypes(typeof(T), out var keyType, out var valueType)
                ? (DictionaryRule)GenericTypes.Create(typeof(DictionaryRule<,>), keyType, valueType)
                : s_noDictionary;
            return ReferenceEquals(rule, s_noDictionary) ? null : rule;
        }
    }

    public override bool TryAutoMapMember(MemberInfo member, EmptyValueStrategy emptyValueStrategy, HashSet<Type>? typeStack, [NotNullWhen(true)] out IMap? map)
    {
        var columns = MemberMapper.GetDefaultColumns(member);
        var cellsColumns = MemberMapper.GetDefaultCellsColumns(member) ?? MemberMapper.GetDefaultSplitColumns(member, columns);
        return AutoMapper.TryCreateMap<T>(member, columns, cellsColumns, emptyValueStrategy, isAutoMapping: true, typeStack, out map);
    }

    public override ExcelClassMap GetOrCreateNestedMap(IMap parentMap, object? context, EmptyValueStrategy emptyValueStrategy)
    {
        if (AutoMapper.GetExistingMap(parentMap, context) is { } existingMap)
        {
            return existingMap as ExcelClassMap<T>
                ?? throw new InvalidOperationException($"Expression is already mapped differently as {existingMap.GetType()}.");
        }

        // Nested objects mapped through their members aren't auto-mapped.
        var classMap = new ExcelClassMap<T>(emptyValueStrategy);
        AutoMapper.AddExistingMap(parentMap, context, classMap);
        return classMap;
    }

    /// <summary>
    /// Adds the mappers of the type's built-in conversion, if it has one.
    /// </summary>
    public bool TryAddDefaultMappers(IList<ICellMapper<T>> mappers)
    {
        var addDefaultMappers = _addDefaultMappers ??= FindDefaultMappers() ?? s_noDefaultMappers;
        if (ReferenceEquals(addDefaultMappers, s_noDefaultMappers))
        {
            return false;
        }

        addDefaultMappers(mappers);
        return true;
    }

    // The built-in conversion for T, or for the underlying type of a nullable T.
    private static Action<IList<ICellMapper<T>>>? FindDefaultMappers()
    {
        var type = typeof(T).GetNullableTypeOrThis(out _);

        // Mappers hold no state (formats and other options come with the cell), so the mappers
        // found here serve every member of the type, and of its nullable version.
        object? mapper =
            // object and IConvertible members receive the cell's typed value (a double, DateTime,
            // bool, string, etc.), not its text.
            type == typeof(object) || type == typeof(IConvertible) ? new ObjectMapper<T>() :
            type == typeof(string) ? StringMapper.Instance :
            type == typeof(bool) ? BoolMapper.Instance :
            type == typeof(Guid) ? GuidMapper.Instance :
            type == typeof(Version) ? VersionMapper.Instance :
            type == typeof(DateTime) ? DateTimeMapper.Instance :
            type == typeof(DateTimeOffset) ? DateTimeOffsetMapper.Instance :
            type == typeof(TimeSpan) ? TimeSpanMapper.Instance :
#if NET6_0_OR_GREATER
            type == typeof(DateOnly) ? DateOnlyMapper.Instance :
            type == typeof(TimeOnly) ? TimeOnlyMapper.Instance :
#endif
            type == typeof(Uri) ? UriMapper.Instance :
            null;

        if (type.IsEnum)
        {
            // Values with a DescriptionAttribute are also mapped from their description.
            var enumMapper = (ICellMapper<T>)EnumMappers.Create(type);
            var descriptions = GetEnumDescriptions(type);
            var descriptionMapper = descriptions is null ? null : new MappingDictionaryMapper<T>(descriptions, null, MappingDictionaryBehavior.Optional);
            return mappers =>
            {
                mappers.Add(enumMapper);
                if (descriptionMapper is not null)
                {
                    mappers.Add(descriptionMapper);
                }
            };
        }

        if (mapper is null && AutoMapper.CanConstructObject(type))
        {
#if NET7_0_OR_GREATER
            // Numbers. Number value types (all the built-in ones) also convert nullable members
            // without boxing.
            if (CreateBuiltInNumberMapper(type) is { } builtInMapper)
            {
                mapper = builtInMapper;
            }
            else if (type.ImplementsGenericInterface(typeof(INumberBase<>), out var numberType))
            {
                mapper = GenericTypes.Create(numberType.IsValueType ? typeof(StructNumberMapper<>) : typeof(INumberBaseMapper<>), numberType);
            }
#else
            // The netstandard builds don't have INumberBase<T>, so built-in numeric types use a
            // mapper that behaves the same way (including number styles and format providers).
            if (NumberMapper.IsSupportedType(type))
            {
                mapper = NumberMapper.Create(type);
            }
#endif
            else if (type.ImplementsInterface(typeof(IConvertible)))
            {
                mapper = new ChangeTypeMapper<T>();
            }
#if NET7_0_OR_GREATER
            else if (type.ImplementsGenericInterface(typeof(IParsable<>), out var parsableType))
            {
                mapper = GenericTypes.Create(parsableType.IsValueType ? typeof(StructParsableMapper<>) : typeof(ParsableMapper<>), parsableType);
            }
#endif
        }

        if (mapper is null)
        {
            return null;
        }

        var typeMapper = (ICellMapper<T>)mapper;
        return mappers => mappers.Add(typeMapper);
    }

#if NET7_0_OR_GREATER
    // The mappers of the built-in number types, created directly so that native AOT
    // has their code. Other number types (a user's INumberBase<T>) are found by their interface.
    private static object? CreateBuiltInNumberMapper(Type type)
        => type == typeof(int) ? new StructNumberMapper<int>() :
            type == typeof(long) ? new StructNumberMapper<long>() :
            type == typeof(double) ? new StructNumberMapper<double>() :
            type == typeof(decimal) ? new StructNumberMapper<decimal>() :
            type == typeof(float) ? new StructNumberMapper<float>() :
            type == typeof(short) ? new StructNumberMapper<short>() :
            type == typeof(byte) ? new StructNumberMapper<byte>() :
            type == typeof(sbyte) ? new StructNumberMapper<sbyte>() :
            type == typeof(ushort) ? new StructNumberMapper<ushort>() :
            type == typeof(uint) ? new StructNumberMapper<uint>() :
            type == typeof(ulong) ? new StructNumberMapper<ulong>() :
            type == typeof(nint) ? new StructNumberMapper<nint>() :
            type == typeof(nuint) ? new StructNumberMapper<nuint>() :
            type == typeof(Half) ? new StructNumberMapper<Half>() :
            type == typeof(Int128) ? new StructNumberMapper<Int128>() :
            type == typeof(UInt128) ? new StructNumberMapper<UInt128>() :
            type == typeof(BigInteger) ? new StructNumberMapper<BigInteger>() :
            type == typeof(Complex) ? new StructNumberMapper<Complex>() :
            null;
#endif

    private static Dictionary<string, T>? GetEnumDescriptions(Type enumType)
    {
        var descriptions = enumType.GetFields(BindingFlags.Public | BindingFlags.Static)
            .Select(f => (Field: f, Description: f.GetCustomAttribute<DescriptionAttribute>()?.Description))
            .Where(x => !string.IsNullOrEmpty(x.Description))
            .ToDictionary(x => x.Description!, x => (T)x.Field.GetValue(null)!);
        return descriptions.Count > 0 ? descriptions : null;
    }
}

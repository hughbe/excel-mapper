using System;
using System.Collections;
using System.Collections.Frozen;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Collections.ObjectModel;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Reflection;
using System.Threading;
using ExcelDataReader;
using ExcelMapper.Abstractions;
using ExcelMapper.Factories;
using ExcelMapper.Fallbacks;
using ExcelMapper.Mappers;
using ExcelMapper.Readers;

namespace ExcelMapper.Utilities;

public static class AutoMapper
{
    internal static OneToOneMap<T>? CreateOneToOneMap<T>(MemberInfo? member, ICellReaderFactory defaultCellReaderFactory, FallbackStrategy emptyValueStrategy, bool isAutoMapping)
    {
        var map = new OneToOneMap<T>(defaultCellReaderFactory);
        if (!TrySetupMap<T>(member, map, emptyValueStrategy, isAutoMapping))
        {
            return null;
        }

        return map;
    }

    private static bool TrySetupMap<T>(MemberInfo? member, IToOneMap map, FallbackStrategy emptyValueStrategy, bool isAutoMapping)
    {
        // Try to get a well-known mapper for the type.
        // If we are auto mapping and no well-known mapper is found, fail the auto mapping.
        // But, allow `Map(o => o.Value)` as the user can define a custom converter after
        // creating the map.
        // If the user doesn't add any mappers to the property, an ExcelMappingException
        // will be thrown at the point of registering the class map.
        if (TryGetWellKnownMapper<T>(out var mapper))
        {
            map.Pipeline.Mappers.Add(mapper);
        }
        else if (isAutoMapping)
        {
            return false;
        }

        // Setup the fallbacks.
        if (member != null && member.GetCustomAttribute<ExcelDefaultValueAttribute>() is { } defaultValueAttribute)
        {
            map.Pipeline.EmptyFallback = new FixedValueFallback(defaultValueAttribute.Value);
        }
        else
        {
            map.Pipeline.EmptyFallback = CreateEmptyFallback<T>(emptyValueStrategy);
        }

        map.Pipeline.InvalidFallback = s_throwFallback;

        // Apply member attributes.
        if (member != null)
        {
            if (Attribute.IsDefined(member, typeof(ExcelOptionalAttribute)))
            {
                map.Optional = true;
            }
            if (Attribute.IsDefined(member, typeof(ExcelPreserveFormattingAttribute)))
            {
                map.PreserveFormatting = true;
            }
        }

        return true;
    }

    // Cached singleton instances for stateless objects
    private static readonly ICellMapper s_dateTimeMapper = new DateTimeMapper();
    private static readonly ICellMapper s_dateTimeOffsetMapper = new DateTimeOffsetMapper();
    private static readonly ICellMapper s_timeSpanMapper = new TimeSpanMapper();
    private static readonly ICellMapper s_dateOnlyMapper = new DateOnlyMapper();
    private static readonly ICellMapper s_timeOnlyMapper = new TimeOnlyMapper();
    private static readonly ICellMapper s_guidMapper = new GuidMapper();
    private static readonly ICellMapper s_boolMapper = new BoolMapper();
    private static readonly ICellMapper s_stringMapper = new StringMapper();
    private static readonly ICellMapper s_uriMapper = new UriMapper();
    private static readonly ICellMapper s_versionMapper = new VersionMapper();
    private static readonly AllColumnNamesReaderFactory s_allColumnNamesReaderFactory = new();
    private static readonly IFallbackItem s_throwFallback = new ThrowFallback();
    private static readonly IFallbackItem s_nullFallback = new FixedValueFallback(null);

    private static readonly FrozenDictionary<Type, ICellMapper> s_wellKnownTypeMappers = new Dictionary<Type, ICellMapper>
    {
        [typeof(DateTime)] = s_dateTimeMapper,
        [typeof(DateTimeOffset)] = s_dateTimeOffsetMapper,
        [typeof(TimeSpan)] = s_timeSpanMapper,
        [typeof(DateOnly)] = s_dateOnlyMapper,
        [typeof(TimeOnly)] = s_timeOnlyMapper,
        [typeof(Guid)] = s_guidMapper,
        [typeof(bool)] = s_boolMapper,
        [typeof(string)] = s_stringMapper,
        [typeof(object)] = s_stringMapper,
        [typeof(IConvertible)] = s_stringMapper,
        [typeof(Uri)] = s_uriMapper,
        [typeof(Version)] = s_versionMapper,
    }.ToFrozenDictionary();

    private static bool TryGetWellKnownMapper<T>([NotNullWhen(true)] out ICellMapper? mapper)
    {
        var type = typeof(T).GetNullableTypeOrThis(out _);

        // Fast path: Check dictionary for well-known types.
        if (s_wellKnownTypeMappers.TryGetValue(type, out var cachedMapper))
        {
            mapper = cachedMapper;
            return true;
        }
        // Check for enum types.
        else if (type.IsEnum)
        {
            mapper = new EnumMapper(type);
            return true;
        }
        // Check for types implementing interfaces.
        else if (CanConstructObject(type))
        {
            // Check for types implementing IConvertible.
            if (type.ImplementsInterface(typeof(IConvertible)))
            {
                mapper = new ChangeTypeMapper(type);
                return true;
            }
            // Check for types implementing IParsable<T>.
            else if (type.ImplementsGenericInterface(typeof(IParsable<>), out var parsableInterfaceType))
            {
                mapper = (ICellMapper)Activator.CreateInstance(typeof(ParsableMapper<>).MakeGenericType(parsableInterfaceType))!;
                return true;
            }
        }

        mapper = null;
        return false;
    }

    private static IFallbackItem CreateEmptyFallback<T>(FallbackStrategy emptyValueStrategy)
    {
        var isNullable = typeof(T).IsNullable();

        // Empty nullable values should be set to null.
        if (isNullable || !typeof(T).IsValueType)
        {
            return s_nullFallback;
        }
        else if (emptyValueStrategy == FallbackStrategy.SetToDefaultValue)
        {
            return new FixedValueFallbackFactory(() => Activator.CreateInstance(typeof(T)));
        }

        // Throw if we can't set to null or default value.
        return s_throwFallback;
    }
    
    private static readonly Lazy<MethodInfo> s_tryCreateSplitGenericMapMethod = new(
        () => typeof(AutoMapper).GetMethod(nameof(TryCreateSplitMapGeneric), BindingFlags.NonPublic | BindingFlags.Static)!,
        LazyThreadSafetyMode.PublicationOnly);
    private static MethodInfo TryCreateSplitGenericMapMethod => s_tryCreateSplitGenericMapMethod.Value;

    private static bool TryCreateSplitMap(MemberInfo? member, Type listType, ICellsReaderFactory defaultCellsReaderFactory, FallbackStrategy emptyValueStrategy, [NotNullWhen(true)] out IMap? map)
    {
        var elementType = listType.GetElementTypeOrEnumerableType();
        if (elementType == null)
        {
            map = null;
            return false;
        }

        var method = TryCreateSplitGenericMapMethod.MakeGenericMethod([elementType]);
        var parameters = new object?[] { member, listType, defaultCellsReaderFactory, emptyValueStrategy, null };
        var result = (bool)method.InvokeUnwrapped(null, parameters)!;
        if (result)
        {
            map = (IMap)parameters[^1]!;
            return true;
        }

        map = null;
        return false;
    }

    internal static bool TryCreateSplitMapGeneric<TElement>(MemberInfo? member, Type listType, ICellsReaderFactory defaultCellsReaderFactory, FallbackStrategy emptyValueStrategy, [NotNullWhen(true)] out ManyToOneEnumerableMap<TElement>? map)
    {
        // Find the right way of creating the collection and adding items to it.
        if (!TryGetCreateEnumerableFactory<TElement>(listType, out var factory))
        {
            map = null;
            return false;
        }

        // First, get the pipeline for the element. This is used to convert individual values
        // to be added to/included in the collection.
        map = new ManyToOneEnumerableMap<TElement>(defaultCellsReaderFactory, factory);
        if (!TrySetupMap<TElement>(member, map, emptyValueStrategy, isAutoMapping: true))
        {
            map = null;
            return false;
        }

        return true;
    }

    private static bool TryGetCreateEnumerableFactory<TElement>(Type listType, [NotNullWhen(true)] out IEnumerableFactory<TElement>? result)
    {
        // Try well-known collection types
        if (TryGetWellKnownEnumerableFactory(listType, out result))
        {
            return true;
        }

        // Try interface types
        if (listType.IsInterface)
        {
            return TryGetInterfaceEnumerableFactory(listType, out result);
        }

        // Try concrete types
        if (!listType.IsAbstract)
        {
            return TryGetConcreteTypeEnumerableFactory(listType, out result);
        }

        result = default;
        return false;
    }

    private static bool TryGetWellKnownEnumerableFactory<TElement>(Type listType, [NotNullWhen(true)] out IEnumerableFactory<TElement>? result)
    {
        if (listType == typeof(Array) || (listType.IsArray && listType.GetArrayRank() == 1))
        {
            result = new ArrayEnumerableFactory<TElement>();
            return true;
        }
        if (listType == typeof(ImmutableArray<TElement>))
        {
            result = new ImmutableArrayEnumerableFactory<TElement>();
            return true;
        }
        if (listType == typeof(ImmutableList<TElement>) || listType == typeof(IImmutableList<TElement>))
        {
            result = new ImmutableListEnumerableFactory<TElement>();
            return true;
        }
        if (listType == typeof(ImmutableStack<TElement>) || listType == typeof(IImmutableStack<TElement>))
        {
            result = new ImmutableStackEnumerableFactory<TElement>();
            return true;
        }
        if (listType == typeof(ImmutableQueue<TElement>) || listType == typeof(IImmutableQueue<TElement>))
        {
            result = new ImmutableQueueEnumerableFactory<TElement>();
            return true;
        }
        if (listType == typeof(ImmutableSortedSet<TElement>))
        {
            result = new ImmutableSortedSetEnumerableFactory<TElement>();
            return true;
        }
        if (listType == typeof(ImmutableHashSet<TElement>) || listType == typeof(IImmutableSet<TElement>))
        {
            result = new ImmutableHashSetEnumerableFactory<TElement>();
            return true;
        }
        if (listType == typeof(FrozenSet<TElement>))
        {
            result = new FrozenSetEnumerableFactory<TElement>();
            return true;
        }

        result = default;
        return false;
    }

    private static bool TryGetInterfaceEnumerableFactory<TElement>(Type listType, [NotNullWhen(true)] out IEnumerableFactory<TElement>? result)
    {
        // Add values by creating a list and assigning to the property.
        if (typeof(List<TElement>).IsAssignableTo(listType))
        {
            result = new ListEnumerableFactory<TElement>();
            return true;
        }
        if (typeof(HashSet<TElement>).IsAssignableTo(listType))
        {
            result = new HashSetEnumerableFactory<TElement>();
            return true;
        }

        result = default;
        return false;
    }

    private static bool TryGetConcreteTypeEnumerableFactory<TElement>(Type listType, [NotNullWhen(true)] out IEnumerableFactory<TElement>? result)
    {
        var hasDefaultConstructor = listType.GetConstructor([]) is not null;

        // Try default constructor with collection interfaces
        if (hasDefaultConstructor && TryGetDefaultConstructorFactory(listType, out result))
        {
            return true;
        }

        // Try constructor-based factories
        if (TryGetConstructorBasedFactory(listType, out result))
        {
            return true;
        }

        // Try Add method factory
        if (hasDefaultConstructor)
        {
            var addMethod = listType.GetMethod("Add", [typeof(TElement)]);
            if (addMethod != null)
            {
                result = new AddEnumerableFactory<TElement>(listType);
                return true;
            }
        }

        // Try ObservableCollection constructor
        var ctor = listType.GetConstructor([typeof(ObservableCollection<TElement>)]);
        if (ctor != null)
        {
            result = new ReadOnlyObservableCollectionEnumerableFactory<TElement>(listType);
            return true;
        }

        result = default;
        return false;
    }

    private static bool TryGetDefaultConstructorFactory<TElement>(Type listType, [NotNullWhen(true)] out IEnumerableFactory<TElement>? result)
    {
        // Add values through IList<TElement>.Add(TElement item).
        if (listType.ImplementsInterface(typeof(IList<TElement>)))
        {
            result = new IListTImplementingEnumerableFactory<TElement>(listType);
            return true;
        }

        // Add values through ISet<TElement>.Add(TElement item).
        if (listType.ImplementsInterface(typeof(ISet<TElement>)))
        {
            result = new ISetTImplementingEnumerableFactory<TElement>(listType);
            return true;
        }

        // Add values through ICollection<TElement>.Add(TElement item).
        if (listType.ImplementsInterface(typeof(ICollection<TElement>)))
        {
            result = new ICollectionTImplementingEnumerableFactory<TElement>(listType);
            return true;
        }

        // Add values through IList.Add(TElement item).
        if (listType.ImplementsInterface(typeof(IList)))
        {
            result = new IListImplementingEnumerableFactory<TElement>(listType);
            return true;
        }

        result = default;
        return false;
    }

    private static bool TryGetConstructorBasedFactory<TElement>(Type listType, [NotNullWhen(true)] out IEnumerableFactory<TElement>? result)
    {
        // Check if the type has .ctor(ISet<T>) such as ReadOnlySet.
        var ctor = listType.GetConstructor([typeof(ISet<TElement>)]);
        if (ctor != null)
        {
            result = new ConstructorSetEnumerableFactory<TElement>(listType);
            return true;
        }

        // Check if the type has .ctor(IList<T>) .ctor(ICollection) or .ctor(IEnumerable<T>) such as ReadOnlyCollection, Queue or Stack.
        ctor = listType.GetConstructor([typeof(IList<TElement>)])
            ?? listType.GetConstructor([typeof(IEnumerable<TElement>)])
            ?? listType.GetConstructor([typeof(ICollection)]);
        if (ctor != null)
        {
            result = new ConstructorEnumerableFactory<TElement>(listType);
            return true;
        }

        result = default;
        return false;
    }

    public static ExcelClassMap GetOrCreateNestedMap(IMap parentMap, Type memberType, object? context, FallbackStrategy emptyValueStrategy)
    {
        var method = GetOrCreateNestedMapGenericMethod.MakeGenericMethod(memberType);
        var parameters = new object?[] { parentMap, context, emptyValueStrategy };
        return (ExcelClassMap)method.InvokeUnwrapped(null, parameters)!;
    }

    private static IMap GetOrCreateNestedMapGeneric<TProperty>(IMap parentMap, object? context, FallbackStrategy emptyValueStrategy)
    {
        if (GetExistingMap(parentMap, context) is { } existingMap)
        {
            if (existingMap is not ExcelClassMap<TProperty> existingTypeMap)
            {
                throw new InvalidOperationException($"Expression is already mapped differently as {existingMap.GetType()}.");
            }

            return existingTypeMap;
        }

        // By default, do not auto map nested fields.
        var classMap = new ExcelClassMap<TProperty>(emptyValueStrategy);
        AddExistingMap(parentMap, context, classMap);
        return classMap;
    }

    private static readonly Lazy<MethodInfo> s_getOrCreateArrayIndexerMapGenericMethod = new(
        () => typeof(AutoMapper).GetMethod(nameof(GetOrCreateArrayIndexerMapGeneric), BindingFlags.NonPublic | BindingFlags.Static)!,
        LazyThreadSafetyMode.PublicationOnly);
    private static MethodInfo GetOrCreateArrayIndexerMapGenericMethod => s_getOrCreateArrayIndexerMapGenericMethod.Value;

    internal static IMap GetOrCreateArrayIndexerMap(IMap parentMap, Type arrayType, object? context, Type elementType)
    {
        var method = GetOrCreateArrayIndexerMapGenericMethod.MakeGenericMethod([elementType]);
        var parameters = new object?[] { parentMap, arrayType, context };
        return (IMap)method.InvokeUnwrapped(null, parameters)!;
    }

    private static ManyToOneEnumerableIndexerMapT<TElement> GetOrCreateArrayIndexerMapGeneric<TElement>(IMap parentMap, Type arrayType, object? context)
    {
        if (GetExistingMap(parentMap, context) is { } existingMap)
        {
            if (existingMap is not ManyToOneEnumerableIndexerMapT<TElement> arrayIndexerMap)
            {
                throw new InvalidOperationException($"Expression is already mapped differently as {existingMap.GetType()}.");
            }

            return arrayIndexerMap;
        }

        // Create a new array indexer map.
        if (!TryCreateArrayIndexerMapGeneric<TElement>(arrayType, GetEmptyValueStrategy(parentMap), out var mapObj))
        {
            throw new ExcelMappingException($"Could not map array of type \"{arrayType}\".");
        }

        AddExistingMap(parentMap, context, mapObj);
        return mapObj;
    }

    private static bool TryCreateArrayIndexerMapGeneric<TElement>(Type listType, FallbackStrategy emptyValueStrategy, [NotNullWhen(true)] out ManyToOneEnumerableIndexerMapT<TElement>? map)
    {
        if (!TryGetCreateEnumerableFactory<TElement>(listType, out var factory))
        {
            map = null;
            return false;
        }

        map = new ManyToOneEnumerableIndexerMapT<TElement>(factory);
        return true;
    }

    private static readonly Lazy<MethodInfo> s_getOrCreateMultidimensionalIndexerMapGenericMethod = new(
        () => typeof(AutoMapper).GetMethod(nameof(GetOrCreateMultidimensionalIndexerMapGeneric), BindingFlags.NonPublic | BindingFlags.Static)!,
        LazyThreadSafetyMode.PublicationOnly);
    private static MethodInfo GetOrCreateMultidimensionalIndexerMapGenericMethod => s_getOrCreateMultidimensionalIndexerMapGenericMethod.Value;

    internal static IMap GetOrCreateMultidimensionalIndexerMap(IMap parentMap, Type arrayType, Type elementType, object? context)
    {
        var method = GetOrCreateMultidimensionalIndexerMapGenericMethod.MakeGenericMethod([elementType]);
        var parameters = new object?[] { parentMap, arrayType, context };
        return (IMap)method.InvokeUnwrapped(null, parameters)!;
    }

    private static ManyToOneMultidimensionalIndexerMapT<TElement> GetOrCreateMultidimensionalIndexerMapGeneric<TElement>(IMap parentMap, Type arrayType, object? context)
    {
        if (GetExistingMap(parentMap, context) is { } existingMap)
        {
            if (existingMap is not ManyToOneMultidimensionalIndexerMapT<TElement> multidimensionalMap)
            {
                throw new InvalidOperationException($"Expression is already mapped differently as {existingMap.GetType()}.");
            }

            return multidimensionalMap;
        }

        // Create a new array indexer map.
        var factory = new MultidimensionalArrayFactory<TElement>();
        var map = new ManyToOneMultidimensionalIndexerMapT<TElement>(factory);
        AddExistingMap(parentMap, context, map);
        return map;
    }

    private static readonly Lazy<MethodInfo> s_getOrCreateDictionaryIndexerMapGenericMethod = new(
        () => typeof(AutoMapper).GetMethod(nameof(GetOrCreateDictionaryIndexerMapGeneric), BindingFlags.NonPublic | BindingFlags.Static)!,
        LazyThreadSafetyMode.PublicationOnly);
    private static MethodInfo GetOrCreateDictionaryIndexerMapGenericMethod => s_getOrCreateDictionaryIndexerMapGenericMethod.Value;

    internal static IDictionaryIndexerMap GetOrCreateDictionaryIndexerMap(IMap parentMap, Type dictionaryType, Type keyType, Type valueType, object? context)
    {
        var method = GetOrCreateDictionaryIndexerMapGenericMethod.MakeGenericMethod([keyType, valueType]);
        var parameters = new object?[] { parentMap, dictionaryType, context };
        return (IDictionaryIndexerMap)method.InvokeUnwrapped(null, parameters)!;
    }

    private static ManyToOneDictionaryIndexerMapT<TKey, TValue> GetOrCreateDictionaryIndexerMapGeneric<TKey, TValue>(IMap parentMap, Type dictionaryType, object? context) where TKey : notnull
    {
        if (GetExistingMap(parentMap, context) is { } existingMap)
        {
            if (existingMap is not ManyToOneDictionaryIndexerMapT<TKey, TValue> dictionaryIndexerMap)
            {
                throw new InvalidOperationException($"Expression is already mapped differently as {existingMap.GetType()}.");
            }

            return dictionaryIndexerMap;
        }

        // Create a new dictionary indexer map.
        if (!TryCreateDictionaryIndexerMapGeneric<TKey, TValue>(dictionaryType, GetEmptyValueStrategy(parentMap), out var mapObj))
        {
            throw new ExcelMappingException($"Could not map dictionary of type \"{dictionaryType}\".");
        }

        AddExistingMap(parentMap, context, mapObj);
        return mapObj;
    }

    private static bool TryCreateDictionaryIndexerMapGeneric<TKey, TValue>(Type dictionaryType, FallbackStrategy emptyValueStrategy, [NotNullWhen(true)] out ManyToOneDictionaryIndexerMapT<TKey, TValue>? map) where TKey : notnull
    {
        if (!TryGetCreateDictionaryFactory<TKey, TValue>(dictionaryType, out var factory))
        {
            map = null;
            return false;
        }

        map = new ManyToOneDictionaryIndexerMapT<TKey, TValue>(factory);
        return true;
    }

    private static readonly Lazy<MethodInfo> s_tryCreateDictionaryMapGeneric = new(
        () => typeof(AutoMapper).GetMethod(nameof(TryCreateGenericDictionaryMap), BindingFlags.NonPublic | BindingFlags.Static)!,
        LazyThreadSafetyMode.PublicationOnly);
    private static MethodInfo TryCreateDictionaryMapGeneric => s_tryCreateDictionaryMapGeneric.Value;

    private static bool TryCreateDictionaryMap<TMember>(MemberInfo? member, FallbackStrategy emptyValueStrategy, bool isAutoMapping, [NotNullWhen(true)] out IMap? map)
    {
        if (!TryGetDictionaryKeyValueType(typeof(TMember), out var keyType, out var valueType))
        {
            map = null;
            return false;
        }

        var method = TryCreateDictionaryMapGeneric.MakeGenericMethod(keyType, valueType);
        var parameters = new object?[] { member, typeof(TMember), emptyValueStrategy, isAutoMapping, null };
        var result = (bool)method.InvokeUnwrapped(null, parameters)!;
        if (result)
        {
            map = (IMap)parameters[^1]!;
            return true;
        }

        map = null;
        return false;
    }

    internal static bool TryGetDictionaryKeyValueType(Type dictionaryType, [NotNullWhen(true)] out Type? keyType, [NotNullWhen(true)] out Type? valueType)
    {
        // We should be able to parse anything that implements IEnumerable<KeyValuePair<TKey, TValue>>
        if (dictionaryType.ImplementsGenericInterface(typeof(IEnumerable<>), out Type? keyValuePairType))
        {
            if (keyValuePairType.IsGenericType && keyValuePairType.GetGenericTypeDefinition() == typeof(KeyValuePair<,>))
            {
                Type[] arguments = keyValuePairType.GenericTypeArguments;
                keyType = arguments[0];
                valueType = arguments[1];
                return true;
            }
        }

        // Parse regular IDictionary before checking for Add methods.
        // IDictionary also has Add(object, object) but we want to use string keys by convention.
        if (dictionaryType == typeof(IDictionary) || dictionaryType.ImplementsInterface(typeof(IDictionary)))
        {
            keyType = typeof(string);
            valueType = typeof(object);
            return true;
        }

        // Check for types like StringDictionary that have Add(TKey, TValue) method and implement IEnumerable.
        // This handles dictionary-like types that don't implement IDictionary<TKey, TValue> or IDictionary.
        if (dictionaryType.ImplementsInterface(typeof(IEnumerable)))
        {
            // If there are multiple add methods, we can't determine the key/value types.
            var addMethods = dictionaryType.GetMethods()
                .Where(m => m.Name == "Add" && m.GetParameters().Length == 2)
                .ToArray();
            
            if (addMethods.Length == 1)
            {
                var parameters = addMethods[0].GetParameters();
                keyType = parameters[0].ParameterType;
                valueType = parameters[1].ParameterType;
                return true;
            }
        }
        
        keyType = null;
        valueType = null;
        return false;
    }

    internal static bool TryCreateGenericDictionaryMap<TKey, TValue>(MemberInfo? member, Type dictionaryType, FallbackStrategy emptyValueStrategy, bool isAutoMapping, [NotNullWhen(true)] out ManyToOneDictionaryMap<TKey, TValue>? map) where TKey : notnull
    {
        // Find the right way of creating the dictionary and adding items to it.
        if (!TryGetCreateDictionaryFactory<TKey, TValue>(dictionaryType, out var factory))
        {
            map = null;
            return false;
        }

        // Default to all columns.
        var defaultReaderFactory = MemberMapper.GetDefaultCellsReaderFactory(member) ?? s_allColumnNamesReaderFactory;
        map = new ManyToOneDictionaryMap<TKey, TValue>(defaultReaderFactory, factory);
        if (!TrySetupMap<TValue>(member, map, emptyValueStrategy, isAutoMapping))
        {
            map = null;
            return false;
        }

        return true;
    }

    private static bool TryGetCreateDictionaryFactory<TKey, TValue>(Type dictionaryType, [NotNullWhen(true)] out IDictionaryFactory<TKey, TValue>? result) where TKey : notnull
    {
        if (dictionaryType == typeof(ImmutableDictionary<TKey, TValue>) || dictionaryType == typeof(IImmutableDictionary<TKey, TValue>))
        {
            result = new ImmutableDictionaryFactory<TKey, TValue>();
            return true;
        }
        else if (dictionaryType == typeof(ImmutableSortedDictionary<TKey, TValue>))
        {
            result = new ImmutableSortedDictionaryFactory<TKey, TValue>();
            return true;
        }
        else if (dictionaryType == typeof(FrozenDictionary<TKey, TValue>))
        {
            result = new FrozenDictionaryFactory<TKey, TValue>();
            return true;
        }
        else if (dictionaryType.IsInterface)
        {
            if (typeof(Dictionary<TKey, TValue>).IsAssignableTo(dictionaryType))
            {
                result = new DictionaryFactory<TKey, TValue>();
                return true;
            }
        }
        else if (dictionaryType.GetConstructor([]) is not null)
        {
            if (dictionaryType.ImplementsInterface(typeof(IDictionary<TKey, TValue>)))
            {
                result = new IDictionaryTImplementingFactory<TKey, TValue>(dictionaryType);
                return true;
            }
            else if (dictionaryType.ImplementsInterface(typeof(IDictionary)))
            {
                result = new IDictionaryImplementingFactory<TKey, TValue>(dictionaryType);
                return true;
            }

            // Check if the type has Add(TKey, TValue) such as StringDictionary.
            var addMethod = dictionaryType.GetMethod("Add", [typeof(TKey), typeof(TValue)]);
            if (addMethod != null)
            {
                result = new AddDictionaryFactory<TKey, TValue>(dictionaryType);
                return true;
            }
        }
        else
        {
            // Check if the type has .ctor(IDictionary<TKey, TValue>) such as ReadOnlyDictionary<TKey, TValue>.
            var ctor = dictionaryType.GetConstructor([typeof(IDictionary<TKey, TValue>)]);
            if (ctor != null)
            {
                result = new ConstructorDictionaryFactory<TKey, TValue>(dictionaryType);
                return true;
            }
        }

        result = default;
        return false;
    }

    internal static bool TryCreateObjectMap<T>(FallbackStrategy emptyValueStrategy, [NotNullWhen(true)] out ExcelClassMap<T>? classMap)
        => TryCreateObjectMap(emptyValueStrategy, out classMap, null);

    private static bool TryCreateObjectMap<T>(FallbackStrategy emptyValueStrategy, [NotNullWhen(true)] out ExcelClassMap<T>? classMap, HashSet<Type>? typeStack)
    {
        Type type = typeof(T);
        if (!CanConstructObject(type))
        {
            classMap = null;
            return false;
        }

        // Initialize the type stack on first call
        typeStack ??= [];

        // Check for circular reference
        if (!typeStack.Add(type))
        {
            throw new ExcelMappingException($"Circular reference detected: type \"{type.Name}\" references itself through its members. Consider applying the ExcelIgnore attribute to break the cycle.");
        }

        try
        {
            var map = new ExcelClassMap<T>(emptyValueStrategy);
            
            // Process properties
            foreach (var property in type.GetProperties(BindingFlags.Public | BindingFlags.Instance))
            {
                if (!ShouldAutoMapProperty(property))
                {
                    continue;
                }

                if (!TryAutoMapAndAddMember(property, property.PropertyType, map, emptyValueStrategy, typeStack, out classMap))
                {
                    return false;
                }
            }

            // Process fields
            foreach (var field in type.GetFields(BindingFlags.Public | BindingFlags.Instance))
            {
                if (!TryAutoMapAndAddMember(field, field.FieldType, map, emptyValueStrategy, typeStack, out classMap))
                {
                    return false;
                }
            }

            classMap = map;
            return true;
        }
        finally
        {
            // Remove the type from the stack as we exit
            typeStack.Remove(type);
        }
    }
    private static bool TryAutoMapAndAddMember<T>(
        MemberInfo member, 
        Type memberType, 
        ExcelClassMap<T> map, 
        FallbackStrategy emptyValueStrategy, 
        HashSet<Type>? typeStack,
        [NotNullWhen(false)] out ExcelClassMap<T>? classMap)
    {
        // Ignore members with ExcelIgnoreAttribute.
        if (Attribute.IsDefined(member, typeof(ExcelIgnoreAttribute)))
        {
            classMap = null;
            return true; // Continue processing other members
        }

        // Infer the mapping for each member (property/field) belonging to the type.
        var method = TryAutoMapMemberMethod.MakeGenericMethod(memberType);
        var parameters = new object?[] { member, emptyValueStrategy, typeStack, null };
        var result = (bool)method.InvokeUnwrapped(null, parameters)!;
        if (!result)
        {
            classMap = null!;
            return false;
        }

        // Get the out parameter representing the map for the member.
        map.Properties.Add(new ExcelPropertyMap(member, (IMap)parameters[3]!));
        classMap = null;
        return true;
    }

    private static bool CanConstructObject(Type type)
    {
        // Value types can always be constructed.
        if (type.IsValueType)
        {
            return true;
        }

        // Interface and abstract types cannot be constructed.
        if (type.IsInterface)
        {
            return false;
        }
        if (type.IsAbstract)
        {
            return false;
        }

        // Must have a public parameterless constructor.
        if (type.GetConstructor([]) is null)
        {
            return false;
        }

        return true;
    }

    private static bool ShouldAutoMapProperty(PropertyInfo property)
    {
        // Property must have a setter.
        if (!property.CanWrite)
        {
            return false;
        }
        // Property must be a public instance property.
        if (!property.SetMethod!.IsPublic || property.SetMethod.IsStatic)
        {
            return false;
        }

        // Property must not be an indexer.
        if (property.GetIndexParameters().Length > 0)
        {
            return false;
        }

        // Otherwise, this property can be mapped.
        return true;
    }

    private static FallbackStrategy GetEmptyValueStrategy(IMap map)
    {
        if (map is ExcelClassMap classMap)
        {
            return classMap.EmptyValueStrategy;
        }

        return FallbackStrategy.ThrowIfPrimitive;
    }

    private static IMap? GetExistingMap(IMap parentMap, object? context)
    {
        if (parentMap is ExcelClassMap parentClassMap)
        {
            return parentClassMap.Properties.FirstOrDefault(m => m.Member.Equals((MemberInfo)context!))?.Map;
        }
        else if (parentMap is IEnumerableIndexerMap enumerableIndexerMap)
        {
            return enumerableIndexerMap.Values.TryGetValue((int)context!, out var map) ? map : null;
        }
        else if (parentMap is IMultidimensionalIndexerMap multidimensionalIndexerMap)
        {
            // Need to find the key that matches the indices.
            // Cannot use TryGetValue as arrays do not implement equality.
            var indices = (int[])context!;
            foreach (var kvp in multidimensionalIndexerMap.Values)
            {
                if (kvp.Key.SequenceEqual(indices))
                {
                    return kvp.Value;
                }
            }

            return null;
        }
        else
        {
            var dictionaryIndexerMap = (IDictionaryIndexerMap)parentMap;
            return dictionaryIndexerMap.Values.TryGetValue((string)context!, out var map) ? map : null;
        }
    }

    private static IMap AddExistingMap(IMap parentMap, object? context, IMap map)
    {
        if (parentMap is ExcelClassMap parentClassMap)
        {
            parentClassMap.Properties.Add(new ExcelPropertyMap((MemberInfo)context!, map));
        }
        else if (parentMap is IEnumerableIndexerMap enumerableIndexerMap)
        {
            enumerableIndexerMap.Values[(int)context!] = map;
        }
        else if (parentMap is IMultidimensionalIndexerMap multidimensionalIndexerMap)
        {
            multidimensionalIndexerMap.Values[(int[])context!] = map;
        }
        else
        {
            var dictionaryIndexerMap = (IDictionaryIndexerMap)parentMap;
            dictionaryIndexerMap.Values[(string)context!] = map;
        }

        return map;
    }

    private static readonly Lazy<MethodInfo> s_getOrCreateNestedMapGenericMethod = new(
        () => typeof(AutoMapper).GetMethod(nameof(GetOrCreateNestedMapGeneric), BindingFlags.NonPublic | BindingFlags.Static)!,
        LazyThreadSafetyMode.PublicationOnly);
    private static MethodInfo GetOrCreateNestedMapGenericMethod => s_getOrCreateNestedMapGenericMethod.Value;

    /// <summary>
    /// Creates a class map for the given type using the given strategy.
    /// </summary>
    /// <param name="emptyValueStrategy">The default strategy to use when the value of a cell is empty.</param>
    /// <param name="classMap">The class map for the given type.</param>
    /// <returns>True if the class map could be created, else false.</returns>
    public static bool TryCreateClassMap<T>(FallbackStrategy emptyValueStrategy, [NotNullWhen(true)] out ExcelClassMap<T>? result)
    {
        if (!Enum.IsDefined(emptyValueStrategy))
        {
            throw new ArgumentException($"Invalid value \"{emptyValueStrategy}\".", nameof(emptyValueStrategy));
        }

        if (!TryCreateMap<T>(null, typeof(T), new ColumnIndexReaderFactory(0), s_allColumnNamesReaderFactory, emptyValueStrategy, isAutoMapping: true, typeStack: null, out var classMap))
        {
            result = null;
            return false;
        }

        if (classMap is ExcelClassMap<T> typedMap)
        {
            result = typedMap;
            return true;
        }

        // Wrap the class map in a BuiltinClassMap to satisfy the generic type.
        result = new BuiltinClassMap<T>(classMap);
        return true;
    }

    private static readonly Lazy<MethodInfo> s_tryAutoMapMemberMethod = new(
        () => typeof(AutoMapper).GetMethod(nameof(TryAutoMapMember), BindingFlags.NonPublic | BindingFlags.Static)!,
        LazyThreadSafetyMode.PublicationOnly);
    private static MethodInfo TryAutoMapMemberMethod => s_tryAutoMapMemberMethod.Value;

    /// <summary>
    /// Creates a map to assign a value to a specific member.
    /// </summary>
    /// <typeparam name="TMember">The target type.</typeparam>
    /// <param name="member"></param>
    /// <param name="emptyValueStrategy">The behaviour if the value is empty.</param>
    /// <param name="typeStack">Stack of types being processed to detect circular references.</param>
    /// <param name="map">The pipeline.</param>
    /// <returns>True if the member is able to be mapped.</returns>
    private static bool TryAutoMapMember<TMember>(MemberInfo member, FallbackStrategy emptyValueStrategy, HashSet<Type>? typeStack, [NotNullWhen(true)] out IMap? map)
    {
        return TryCreateMap<TMember>(
            member,
            typeof(TMember),
            MemberMapper.GetDefaultCellReaderFactory(member),
            MemberMapper.GetDefaultCellsReaderFactory(member) ?? new CharSplitReaderFactory(MemberMapper.GetDefaultCellReaderFactory(member)),
            emptyValueStrategy,
            isAutoMapping: true,
            typeStack,
            out map);
    }

    private static bool TryCreateMap<T>(MemberInfo? member, Type memberType, ICellReaderFactory defaultCellReaderFactory, ICellsReaderFactory defaultCellsReaderFactory, FallbackStrategy emptyValueStrategy, bool isAutoMapping, HashSet<Type>? typeStack, [NotNullWhen(true)] out IMap? map)
    {
        // Mapping a known type (e.g., string, object, etc.) is supported and simply
        // reads the first column into that value.
        if (CreateOneToOneMap<T>(member, defaultCellReaderFactory, emptyValueStrategy, isAutoMapping) is { } oneToOneMap)
        {
            map = oneToOneMap;
            return true;
        }
        // User may ask to map the row to a dictionary.
        else if (TryCreateDictionaryMap<T>(member, emptyValueStrategy, isAutoMapping, out var dictionaryMap))
        {
            map = dictionaryMap;
            return true;
        }
        // User may ask to map the row to a list.
        else if (TryCreateSplitMap(member, memberType, defaultCellsReaderFactory, emptyValueStrategy, out var enumerableMap))
        {
            map = enumerableMap;
            return true;
        }
        // Otherwise, create the default class map for this type.
        else if (TryCreateObjectMap<T>(emptyValueStrategy, out var classMap, typeStack))
        {
            map = classMap;
            return true;
        }

        map = null;
        return false;
    }

    private class BuiltinClassMap<T> : ExcelClassMap<T>
    {
        private IMap BuiltinMap { get; }

        public BuiltinClassMap(IMap primitiveMap)
        {
            BuiltinMap = primitiveMap;
        }

        public override bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, [NotNullWhen(true)] out object? result)
            => BuiltinMap.TryGetValue(sheet, rowIndex, reader, null, out result);
    }
}

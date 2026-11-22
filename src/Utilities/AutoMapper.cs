using System.Collections;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using System.Linq;
using ExcelMapper.Fallbacks;
using ExcelMapper.Readers;

namespace ExcelMapper.Utilities;

/// <summary>
/// Creates maps from a type and its members' attributes: a value with the type's built-in
/// conversion, a dictionary, a collection, or an object with its own members, tried in that
/// order. What each type supports is decided once, by its <see cref="TypeRule"/>.
/// </summary>
[RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
[RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
internal static class AutoMapper
{
    internal static readonly ICellReaderFactory s_zeroColumnIndexReaderFactory = new ColumnIndexReaderFactory(0);
    internal static readonly ICellsReaderFactory s_allColumnNamesReaderFactory = new AllColumnNamesReaderFactory();

    internal static ValueMap<T>? CreateValueMap<T>(MemberInfo? member, ICellReaderFactory defaultCellReaderFactory, EmptyValueStrategy emptyValueStrategy, bool isAutoMapping)
    {
        var map = new ValueMap<T>(defaultCellReaderFactory);
        if (!TrySetupMap<T>(member, map, emptyValueStrategy, isAutoMapping))
        {
            return null;
        }

        return map;
    }

    internal static bool TrySetupMap<T>(MemberInfo? member, IToOneMap map, EmptyValueStrategy emptyValueStrategy, bool isAutoMapping)
    {
        var pipeline = (ValuePipeline<T>)map.Pipeline;

        // Setup mappers - ExcelConverter, ExcelMappingDictionary, etc.
        var addDefaultMappers = true;
        if (member != null)
        {
            addDefaultMappers = MemberMapper.AddMappers(pipeline, member);
        }

        // If no mappers are speciified, try to get a well-known mapper for the type.
        if (addDefaultMappers)
        {
            // If we are auto mapping and no well-known mapper is found, fail the auto mapping.
            // But, allow `Map(o => o.Value)` as the user can define a custom converter after
            // creating the map.
            // If the user doesn't add any mappers to the property, an ExcelMappingException
            // will be thrown at the point of registering the class map.
            if (!TypeRule<T>.Instance.TryAddDefaultMappers(pipeline.Mappers))
            {
                if (isAutoMapping)
                {
                    return false;
                }
            }
        }

        // Modify the mappers - ExcelFormats, etc.
        if (member != null)
        {
            MemberMapper.ModifyMappers(pipeline, member);
        }

        // Setup the transformers - ExcelTrimString, ExcelTransformer, etc.
        if (member != null)
        {
            MemberMapper.AddTransformers(pipeline, member);
        }

        // Setup the empty fallback - ExcelDefaultValue, ExcelEmptyFallback, etc.
        if (member != null)
        {
            pipeline.EmptyFallback = MemberMapper.GetDefaultEmptyValueFallback<T>(member);
        }

        // If no fallback is specified, use the empty value strategy.
        pipeline.EmptyFallback ??= CreateEmptyFallback<T>(emptyValueStrategy);

        // Setup the invalid fallback - ExcelInvalidValue, ExcelInvalidFallback, etc.
        if (member != null)
        {
            pipeline.InvalidFallback = MemberMapper.GetDefaultInvalidValueFallback<T>(member);
        }
        
        // If no fallback is specified, use the throw fallback.
        pipeline.InvalidFallback ??= ThrowFallback<T>.Instance;

        // Apply member attributes - ExcelOptional, ExcelPreserveFormatting, etc.
        if (member != null)
        {
            MemberMapper.ApplyMemberAttributes(map, member);
        }

        return true;
    }

    private static ICellFallback<T> CreateEmptyFallback<T>(EmptyValueStrategy emptyValueStrategy)
    {
        // Empty nullable values should be set to null.
        if (default(T) is null)
        {
            return new FixedValueFallback<T>(null);
        }
        else if (emptyValueStrategy == EmptyValueStrategy.DefaultValue)
        {
            return new FixedValueFallback<T>(default(T));
        }

        // Throw if we can't set to null or default value.
        return ThrowFallback<T>.Instance;
    }

    /// <summary>
    /// Maps a member (or, with no member, a whole row) of type <typeparamref name="T"/>: as a value
    /// if the type has a built-in conversion or the member's attributes give one, otherwise as a
    /// dictionary, a collection or an object.
    /// </summary>
    internal static bool TryCreateMap<T>(MemberInfo? member, ICellReaderFactory defaultCellReaderFactory, ICellsReaderFactory defaultCellsReaderFactory, EmptyValueStrategy emptyValueStrategy, bool isAutoMapping, HashSet<Type>? typeStack, [NotNullWhen(true)] out IMap? map)
    {
        if (CreateValueMap<T>(member, defaultCellReaderFactory, emptyValueStrategy, isAutoMapping) is { } valueMap)
        {
            map = valueMap;
            return true;
        }

        var rule = TypeRule<T>.Instance;
        if (rule.Dictionary is { } dictionaryRule && dictionaryRule.TryCreateMap(member, typeof(T), emptyValueStrategy, isAutoMapping, out var dictionaryMap))
        {
            map = dictionaryMap;
            return true;
        }

        if (rule.Collection is { } collectionRule && collectionRule.TryCreateMap(member, typeof(T), defaultCellsReaderFactory, emptyValueStrategy, out map))
        {
            return true;
        }

        if (TryCreateObjectMap<T>(emptyValueStrategy, out var classMap, typeStack))
        {
            map = classMap;
            return true;
        }

        map = null;
        return false;
    }

    /// <summary>
    /// Creates a class map for the given type using the given strategy.
    /// </summary>
    /// <param name="emptyValueStrategy">The default strategy to use when the value of a cell is empty.</param>
    /// <param name="result">The class map for the given type.</param>
    /// <returns>True if the class map could be created, else false.</returns>
    public static bool TryCreateClassMap<T>(EmptyValueStrategy emptyValueStrategy, [NotNullWhen(true)] out ExcelClassMap<T>? result)
    {
        EnumUtilities.ValidateIsDefined(emptyValueStrategy);

        if (!TryCreateMap<T>(null, new ColumnIndexReaderFactory(0), s_allColumnNamesReaderFactory, emptyValueStrategy, isAutoMapping: true, typeStack: null, out var objectMap))
        {
            result = null;
            return false;
        }

        if (objectMap is ExcelClassMap<T> typedMap)
        {
            result = typedMap;
            return true;
        }

        // Wrap the class map in a BuiltinClassMap to satisfy the generic type.
        result = new ExcelClassMap<T>(emptyValueStrategy)
        {
            _valueMap = objectMap
        };
        return true;
    }

    internal static bool TryCreateObjectMap<T>(EmptyValueStrategy emptyValueStrategy, [NotNullWhen(true)] out ExcelClassMap<T>? classMap)
        => TryCreateObjectMap(emptyValueStrategy, out classMap, null);

    private static bool TryCreateObjectMap<T>(EmptyValueStrategy emptyValueStrategy, [NotNullWhen(true)] out ExcelClassMap<T>? classMap, HashSet<Type>? typeStack)
    {
        Type type = typeof(T);
        if (!CanConstructObjectMap(type))
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
                if (!ShouldAutoMapProperty(type, property))
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
        EmptyValueStrategy emptyValueStrategy, 
        HashSet<Type>? typeStack,
        [NotNullWhen(false)] out ExcelClassMap<T>? classMap)
    {
        // Ignore members with ExcelIgnoreAttribute.
        if (MemberAttributes.IsDefined<ExcelIgnoreAttribute>(member))
        {
            classMap = null;
            return true; // Continue processing other members
        }

        if (MemberAttributes.IsDefined<ExcelRowNumberAttribute>(member))
        {
            if (!RowNumberMap.IsSupportedType(memberType))
            {
                throw new ExcelMappingException($"Member \"{member.Name}\" has the ExcelRowNumber attribute, but its type \"{memberType}\" is not a numeric type such as int or long.");
            }

            map.Properties.Add(new ExcelPropertyMap(member, new RowNumberMap(memberType)));
            classMap = null;
            return true;
        }

        if (MemberAttributes.IsDefined<ExcelExtraColumnsAttribute>(member))
        {
            if (TypeRule.For(memberType).Dictionary is not { } dictionaryRule ||
                dictionaryRule.KeyType != typeof(string) ||
                !dictionaryRule.TryCreateMap(member, memberType, emptyValueStrategy, isAutoMapping: true, out var extraColumnsMap))
            {
                throw new ExcelMappingException($"Member \"{member.Name}\" has the ExcelExtraColumns attribute, but its type \"{memberType}\" is not a dictionary with string keys.");
            }

            extraColumnsMap.ReaderFactory = new ExtraColumnsReaderFactory(map);
            map.Properties.Add(new ExcelPropertyMap(member, extraColumnsMap));
            classMap = null;
            return true;
        }

        // Infer the mapping for each member (property/field) belonging to the type.
        if (!TypeRule.For(memberType).TryAutoMapMember(member, emptyValueStrategy, typeStack, out var memberMap))
        {
            classMap = null!;
            return false;
        }

        map.Properties.Add(new ExcelPropertyMap(member, memberMap));
        classMap = null;
        return true;
    }

    internal static bool CanConstructObject(Type type)
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

    private static bool ShouldAutoMapProperty(Type type, PropertyInfo property)
    {
        // Property must not be an indexer.
        if (property.GetIndexParameters().Length > 0)
        {
            return false;
        }

        // Property must have a public instance setter...
        if (property.CanWrite && property.SetMethod!.IsPublic && !property.SetMethod.IsStatic)
        {
            return true;
        }

        // ...or be a read-only property that is set through a constructor parameter, for example
        // a positional record or a class with get-only properties.
        return !property.CanWrite
            && property.GetMethod is { IsPublic: true, IsStatic: false }
            && CanUseConstructorParameters(type)
            && ConstructorUtilities.IsConstructorParameter(type, property.Name, property.PropertyType);
    }

    /// <summary>
    /// Gets whether an object map can create instances of the type, either through a parameterless
    /// constructor or a constructor whose parameters match its properties and fields.
    /// </summary>
    private static bool CanConstructObjectMap(Type type)
    {
        if (CanConstructObject(type))
        {
            return true;
        }

        if (type.IsInterface || type.IsAbstract || !CanUseConstructorParameters(type))
        {
            return false;
        }

        return ConstructorUtilities.FindConstructor(type, ConstructorUtilities.GetReadableMembers(type), out _) is not null;
    }

    /// <summary>
    /// Collections (such as Dictionary or BitArray) are mapped by the collection logic, so their
    /// constructor parameters (such as "comparer" or "length") must not be treated as members.
    /// </summary>
    private static bool CanUseConstructorParameters(Type type) => !typeof(IEnumerable).IsAssignableFrom(type);

    internal static IMap? GetExistingMap(IMap parentMap, object? context)
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

    internal static IMap AddExistingMap(IMap parentMap, object? context, IMap map)
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
}

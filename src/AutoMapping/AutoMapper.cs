using System.Collections;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelMapper.Fallbacks;
using ExcelMapper.Readers;

namespace ExcelMapper.AutoMapping;

/// <summary>
/// Creates maps from a type and its members' attributes: a value with the type's built-in
/// conversion, a dictionary, a collection, or an object with its own members, tried in that
/// order. What each type supports is decided once, by its <see cref="TypeRule"/>.
/// </summary>
[RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
[RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
internal static class AutoMapper
{
    // A whole row read as a value reads its first column.
    internal static readonly Columns s_firstColumn = new ColumnIndexColumns(0);

    internal static ValueMap<T>? CreateValueMap<T>(MemberInfo? member, Columns defaultColumns, EmptyValueStrategy emptyValueStrategy, bool isAutoMapping)
    {
        var map = new ValueMap<T>(defaultColumns);
        if (!TrySetupMap<T>(member, map, emptyValueStrategy, isAutoMapping))
        {
            return null;
        }

        return map;
    }

    internal static bool TrySetupMap<T>(MemberInfo? member, IColumnsMap map, EmptyValueStrategy emptyValueStrategy, bool isAutoMapping)
    {
        var pipeline = (CellPipeline<T>)map.Pipeline;

        // Setup mappers - ExcelConverter, ExcelMappingDictionary, etc.
        var addDefaultMappers = true;
        if (member != null)
        {
            addDefaultMappers = MemberMapper.AddMappers(pipeline, member);
        }

        // If no mappers are specified, try to get a well-known mapper for the type.
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
            MemberMapper.ApplyConversionOptions(pipeline, member);
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
    // Keep in sync with the source generator's TypeWalker.Visit, which follows the same order. See "Auto-mapping and the source generator" in CLAUDE.md.
    internal static bool TryCreateMap<T>(MemberInfo? member, Columns defaultColumns, Columns defaultCellsColumns, EmptyValueStrategy emptyValueStrategy, bool isAutoMapping, HashSet<Type>? typeStack, [NotNullWhen(true)] out IMap? map)
    {
        if (CreateValueMap<T>(member, defaultColumns, emptyValueStrategy, isAutoMapping) is { } valueMap)
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

        if (rule.Collection is { } collectionRule && collectionRule.TryCreateMap(member, typeof(T), defaultCellsColumns, emptyValueStrategy, out map))
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

        if (!TryCreateMap<T>(null, s_firstColumn, AllColumns.Instance, emptyValueStrategy, isAutoMapping: true, typeStack: null, out var objectMap))
        {
            result = null;
            return false;
        }

        if (objectMap is ExcelClassMap<T> typedMap)
        {
            result = typedMap;
            return true;
        }

        // A value, collection or dictionary read from the whole row: wrap it in a class map.
        result = new ExcelClassMap<T>(emptyValueStrategy)
        {
            SingleValueMap = objectMap
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
            throw new ExcelMappingException($"Circular reference detected: type {TypeNames.Get(type)} references itself through its members. Consider applying the ExcelIgnore attribute to break the cycle.");
        }

        try
        {
            var map = new ExcelClassMap<T>(emptyValueStrategy);
            if (!TryAddMembers(map, emptyValueStrategy, typeStack))
            {
                classMap = null;
                return false;
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

    /// <summary>
    /// Adds a map for each public property and field of <typeparamref name="T"/> that
    /// <paramref name="map"/> doesn't map yet.
    /// </summary>
    // Keep in sync with the source generator's TypeWalker.GetMappedMembers. See "Auto-mapping and the source generator" in CLAUDE.md.
    private static bool TryAddMembers<T>(ExcelClassMap<T> map, EmptyValueStrategy emptyValueStrategy, HashSet<Type> typeStack)
    {
        var type = typeof(T);
        foreach (var property in type.GetProperties(BindingFlags.Public | BindingFlags.Instance))
        {
            if (ShouldAutoMapProperty(type, property) && !IsMapped(map, property) &&
                !TryAutoMapAndAddMember(property, property.PropertyType, map, emptyValueStrategy, typeStack))
            {
                return false;
            }
        }

        foreach (var field in type.GetFields(BindingFlags.Public | BindingFlags.Instance))
        {
            if (!IsMapped(map, field) &&
                !TryAutoMapAndAddMember(field, field.FieldType, map, emptyValueStrategy, typeStack))
            {
                return false;
            }
        }

        return true;
    }

    private static bool IsMapped(ExcelClassMap map, MemberInfo member)
    {
        // An indexed loop: enumerating the collection through its interface allocates.
        var members = map.Members;
        for (int i = 0; i < members.Count; i++)
        {
            if (members[i].Member.Equals(member))
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Adds what auto-mapping would map to <paramref name="classMap"/>: the members it doesn't map
    /// yet or, for a type read from a row as a single value or collection (such as a string or a
    /// dictionary of the row's columns), that map. See <see cref="ExcelClassMap{T}.AutoMap"/>.
    /// </summary>
    internal static void AutoMap<T>(ExcelClassMap<T> classMap)
    {
        var emptyValueStrategy = classMap.EmptyValueStrategy;
        if (!TryCreateMap<T>(null, s_firstColumn, AllColumns.Instance, emptyValueStrategy, isAutoMapping: true, typeStack: null, out var map))
        {
            throw new ExcelMappingException($"Cannot auto-map type {TypeNames.Get(typeof(T))}.");
        }

        if (map is not ExcelClassMap<T>)
        {
            classMap.SingleValueMap ??= map;
            return;
        }

        // Add the members to the class map itself, rather than copying them from the map
        // created above, as maps such as extra columns refer to the class map they're in. Every
        // member was auto-mapped above, so this can't fail.
        TryAddMembers(classMap, emptyValueStrategy, [typeof(T)]);
    }

    private static bool TryAutoMapAndAddMember<T>(MemberInfo member, Type memberType, ExcelClassMap<T> map, EmptyValueStrategy emptyValueStrategy, HashSet<Type>? typeStack)
    {
        // Ignore members with ExcelIgnoreAttribute.
        if (MemberAttributes.IsDefined<ExcelIgnoreAttribute>(member))
        {
            return true; // Continue processing other members
        }

        if (MemberAttributes.IsDefined<ExcelRowNumberAttribute>(member))
        {
            if (!RowNumberMap.IsSupportedType(memberType))
            {
                throw new ExcelMappingException($"Member \"{member.Name}\" has the ExcelRowNumber attribute, but its type {TypeNames.Get(memberType)} is not a numeric type such as int or long.");
            }

            map.Members.Add(new ClassMapMember(member, new RowNumberMap(memberType)));
            return true;
        }

        if (MemberAttributes.IsDefined<ExcelExtraColumnsAttribute>(member))
        {
            if (TypeRule.For(memberType).Dictionary is not { } dictionaryRule ||
                dictionaryRule.KeyType != typeof(string) ||
                !dictionaryRule.TryCreateMap(member, memberType, emptyValueStrategy, isAutoMapping: true, out var extraColumnsMap))
            {
                throw new ExcelMappingException($"Member \"{member.Name}\" has the ExcelExtraColumns attribute, but its type {TypeNames.Get(memberType)} is not a dictionary with string keys.");
            }

            extraColumnsMap.Settings.Columns = new ExtraColumns(map);
            map.Members.Add(new ClassMapMember(member, extraColumnsMap));
            return true;
        }

        // Infer the mapping for each member (property/field) belonging to the type.
        if (!TypeRule.For(memberType).TryAutoMapMember(member, emptyValueStrategy, typeStack, out var memberMap))
        {
            return false;
        }

        map.Members.Add(new ClassMapMember(member, memberMap));
        return true;
    }

    // Keep in sync with the source generator's TypeWalker.CanConstructObject. See "Auto-mapping and the source generator" in CLAUDE.md.
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
        => ((IParentMap)parentMap).GetChild(context!);

    internal static IMap AddExistingMap(IMap parentMap, object? context, IMap map)
    {
        ((IParentMap)parentMap).SetChild(context!, map);
        return map;
    }
}

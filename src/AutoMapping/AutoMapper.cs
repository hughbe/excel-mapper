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
            return new FixedValueFallback<T>(default!);
        }
        else if (emptyValueStrategy == EmptyValueStrategy.DefaultValue)
        {
            return new FixedValueFallback<T>(default!);
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

        // A dictionary or collection whose values can't be read from a cell, such as
        // Dictionary<string, Address> or List<Address>, isn't then mapped as an object through its
        // own properties (such as Comparer or Capacity).
        if (rule.Dictionary is { } unreadableDictionary && unreadableDictionary.CanCreate(typeof(T)))
        {
            throw ElementsNotReadable(member, typeof(T), unreadableDictionary.ValueType, "values");
        }

        if (rule.Collection is { } unreadableCollection && unreadableCollection.CanCreate(typeof(T)))
        {
            throw ElementsNotReadable(member, typeof(T), unreadableCollection.ElementType, "elements");
        }

        if (TryCreateObjectMap<T>(emptyValueStrategy, out var classMap, typeStack))
        {
            map = classMap;
            return true;
        }

        map = null;
        return false;
    }

    // Elements and values are read from cells, so they must be values: a collection or dictionary
    // can't hold objects mapped member by member from the row.
    private static ExcelMappingException ElementsNotReadable(MemberInfo? member, Type type, Type elementType, string elements)
    {
        var what = member is null ? $"Type {TypeNames.Get(type)}" : $"Member \"{member.Name}\" of type {TypeNames.Get(type)}";
        var fix = member is null
            ? "Read rows of a type whose members are read from the columns instead."
            : $"Give {TypeNames.Get(elementType)} a converter ([ExcelConverter] on the member), or ignore the member with [ExcelIgnore].";
        return new ExcelMappingException($"{what} can't be auto-mapped: its {elements} are {TypeNames.Get(elementType)}, which can't be read from a cell. A collection's elements and a dictionary's values must be values such as strings, numbers, dates or enums. {fix}");
    }

    /// <summary>
    /// The exception for a type that can't be auto-mapped, naming the member that can't be and why.
    /// Only called when auto-mapping has failed, so it can afford to retrace it.
    /// </summary>
    internal static ExcelMappingException CannotAutoMap(Type type)
    {
        var reason = DescribeFailure(type, path: null, []);
        return new ExcelMappingException(reason is null
            ? $"Cannot auto-map type {TypeNames.Get(type)}."
            : $"Cannot auto-map type {TypeNames.Get(type)}: {reason}");
    }

    // Why a type can't be mapped through its members, or null if it isn't known.
    private static string? DescribeFailure(Type type, string? path, HashSet<Type> visited)
    {
        var subject = path is null ? "it" : $"member \"{path}\" of type {TypeNames.Get(type)}";
        if (!CanConstructObjectMap(type))
        {
            return type.IsInterface || type.IsAbstract
                ? $"{subject} is an interface or abstract class, which can't be created."
                : $"{subject} has neither a public parameterless constructor nor a public constructor whose parameters match its members.";
        }

        if (!visited.Add(type))
        {
            return null;
        }

        foreach (var (member, memberType) in GetAutoMappedMembers(type))
        {
            if (MemberAttributes.IsDefined<ExcelIgnoreAttribute>(member)
                || MemberAttributes.IsDefined<ExcelRowNumberAttribute>(member)
                || MemberAttributes.IsDefined<ExcelExtraColumnsAttribute>(member)
                || TypeRule.For(memberType).TryAutoMapMember(member, EmptyValueStrategy.Error, typeStack: null, out _))
            {
                continue;
            }

            var memberPath = path is null ? member.Name : $"{path}.{member.Name}";
            return DescribeFailure(memberType, memberPath, visited)
                ?? $"member \"{memberPath}\" of type {TypeNames.Get(memberType)} can't be mapped.";
        }

        return null;
    }

    // The members auto-mapping maps, with their types: public instance properties with a public
    // setter or set by a constructor, then public instance fields.
    private static IEnumerable<(MemberInfo Member, Type Type)> GetAutoMappedMembers(Type type)
    {
        foreach (var property in type.GetProperties(BindingFlags.Public | BindingFlags.Instance))
        {
            if (ShouldAutoMapProperty(type, property))
            {
                yield return (property, property.PropertyType);
            }
        }

        foreach (var field in type.GetFields(BindingFlags.Public | BindingFlags.Instance))
        {
            yield return (field, field.FieldType);
        }
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
        foreach (var (member, memberType) in GetAutoMappedMembers(typeof(T)))
        {
            if (!IsMapped(map, member) &&
                !TryAutoMapAndAddMember(member, memberType, map, emptyValueStrategy, typeStack))
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
            throw CannotAutoMap(typeof(T));
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

            MemberMapper.ThrowIfColumnAttributes(member, "ExcelRowNumber", "a row number isn't read from a column");

            map.Members.Add(new ClassMapMember(member, new RowNumberMap(memberType)));
            return true;
        }

        if (MemberAttributes.IsDefined<ExcelExtraColumnsAttribute>(member))
        {
            MemberMapper.ThrowIfColumnAttributes(member, "ExcelExtraColumns", "extra columns are the columns no other member reads");
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

using System.Collections;
using System.Collections.Frozen;
using System.Collections.Immutable;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelMapper.Factories;
using ExcelMapper.Readers;

namespace ExcelMapper.Utilities;

/// <summary>
/// The rule for a type with keys and values (a dictionary, or a collection of key-value pairs):
/// how its values are mapped, and how a dictionary of its type is created. See <see cref="TypeRule"/>.
/// </summary>
[RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
[RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
internal abstract class DictionaryRule
{
    public abstract Type KeyType { get; }

    /// <summary>
    /// Maps a member of the dictionary type to a <see cref="DictionaryMap{TKey, TValue}"/> that
    /// reads a value from each of several columns.
    /// </summary>
    public abstract bool TryCreateMap(MemberInfo? member, Type dictionaryType, EmptyValueStrategy emptyValueStrategy, bool isAutoMapping, [NotNullWhen(true)] out IManyToOneMap? map);

    /// <summary>
    /// Gets the map for values mapped one by one, such as <c>Map(p => p.Values["Name"])</c>.
    /// </summary>
    public abstract IMap GetOrCreateIndexerMap(IMap parentMap, Type dictionaryType, object? context);

    /// <summary>
    /// Gets the key and value types of a dictionary type: the type arguments of a collection of
    /// <see cref="KeyValuePair{TKey, TValue}"/>, <c>string</c> and <c>object</c> for an
    /// <see cref="IDictionary"/>, or the parameters of the only two-parameter <c>Add</c> method of a
    /// collection such as <c>StringDictionary</c>.
    /// </summary>
    public static bool TryGetKeyValueTypes(Type dictionaryType, [NotNullWhen(true)] out Type? keyType, [NotNullWhen(true)] out Type? valueType)
    {
        if (dictionaryType.ImplementsGenericInterface(typeof(IEnumerable<>), out var keyValuePairType) &&
            keyValuePairType.IsGenericType && keyValuePairType.GetGenericTypeDefinition() == typeof(KeyValuePair<,>))
        {
            var arguments = keyValuePairType.GenericTypeArguments;
            keyType = arguments[0];
            valueType = arguments[1];
            return true;
        }

        // IDictionary also has Add(object, object), but keys are column names.
        if (dictionaryType == typeof(IDictionary) || dictionaryType.ImplementsInterface(typeof(IDictionary)))
        {
            keyType = typeof(string);
            valueType = typeof(object);
            return true;
        }

        // A collection with a single Add(TKey, TValue) method, such as StringDictionary.
        if (dictionaryType.ImplementsInterface(typeof(IEnumerable)))
        {
            MethodInfo? addMethod = null;
            foreach (var method in dictionaryType.GetMethods())
            {
                if (method.Name == "Add" && method.GetParameters().Length == 2)
                {
                    if (addMethod is not null)
                    {
                        // More than one Add method, so the key and value types are ambiguous.
                        addMethod = null;
                        break;
                    }

                    addMethod = method;
                }
            }

            if (addMethod is not null)
            {
                var parameters = addMethod.GetParameters();
                keyType = parameters[0].ParameterType;
                valueType = parameters[1].ParameterType;
                return true;
            }
        }

        keyType = null;
        valueType = null;
        return false;
    }
}

[RequiresUnreferencedCode(TrimmingConstants.AutoMappingMessage)]
[RequiresDynamicCode(TrimmingConstants.AutoMappingDynamicCodeMessage)]
internal sealed class DictionaryRule<TKey, TValue> : DictionaryRule where TKey : notnull
{
    public override Type KeyType => typeof(TKey);

    public override bool TryCreateMap(MemberInfo? member, Type dictionaryType, EmptyValueStrategy emptyValueStrategy, bool isAutoMapping, [NotNullWhen(true)] out IManyToOneMap? map)
    {
        if (TryCreateDictionaryMap(member, dictionaryType, emptyValueStrategy, isAutoMapping, out var dictionaryMap))
        {
            map = dictionaryMap;
            return true;
        }

        map = null;
        return false;
    }

    /// <summary>
    /// Creates a <see cref="DictionaryMap{TKey, TValue}"/> for a member of type
    /// <paramref name="dictionaryType"/>, with the value type's built-in conversion. Fails if the
    /// dictionary can't be created, or when auto-mapping, if its values have no conversion.
    /// </summary>
    public static bool TryCreateDictionaryMap(MemberInfo? member, Type dictionaryType, EmptyValueStrategy emptyValueStrategy, bool isAutoMapping, [NotNullWhen(true)] out DictionaryMap<TKey, TValue>? map)
    {
        if (!TryGetBuilder(dictionaryType, out var builder))
        {
            map = null;
            return false;
        }

        // A dictionary reads every column by default.
        var columns = MemberMapper.GetDefaultCellsColumns(member) ?? AllColumns.Instance;
        map = new DictionaryMap<TKey, TValue>(columns, builder);
        if (!AutoMapper.TrySetupMap<TValue>(member, map, emptyValueStrategy, isAutoMapping))
        {
            map = null;
            return false;
        }

        return true;
    }

    public override IMap GetOrCreateIndexerMap(IMap parentMap, Type dictionaryType, object? context)
    {
        if (AutoMapper.GetExistingMap(parentMap, context) is { } existingMap)
        {
            return existingMap as DictionaryIndexerMap<TKey, TValue>
                ?? throw new InvalidOperationException($"Expression is already mapped differently as {existingMap.GetType()}.");
        }

        if (!TryGetBuilder(dictionaryType, out var builder))
        {
            throw new ExcelMappingException($"Could not map dictionary of type \"{dictionaryType}\".");
        }

        var map = new DictionaryIndexerMap<TKey, TValue>(builder);
        AutoMapper.AddExistingMap(parentMap, context, map);
        return map;
    }

    /// <summary>
    /// Finds how to build a dictionary of type <paramref name="dictionaryType"/> from its values.
    /// </summary>
    internal static bool TryGetBuilder(Type dictionaryType, [NotNullWhen(true)] out DictionaryBuilder<TKey, TValue>? builder)
    {
        builder = GetBuilder(dictionaryType);
        return builder is not null;
    }

    private static DictionaryBuilder<TKey, TValue>? GetBuilder(Type dictionaryType)
    {
        if (dictionaryType == typeof(Dictionary<TKey, TValue>))
        {
            return DictionaryBuilder<TKey, TValue>.ForDictionary();
        }

        // Immutable and frozen dictionaries are created from a copy of the values.
        if (dictionaryType == typeof(ImmutableDictionary<TKey, TValue>) || dictionaryType == typeof(IImmutableDictionary<TKey, TValue>))
        {
            return DictionaryBuilder<TKey, TValue>.FromDictionary(static values => ImmutableDictionary.CreateRange(values), copiesValues: true);
        }
        if (dictionaryType == typeof(ImmutableSortedDictionary<TKey, TValue>))
        {
            return DictionaryBuilder<TKey, TValue>.FromDictionary(static values => ImmutableSortedDictionary.CreateRange(values), copiesValues: true);
        }
        if (dictionaryType == typeof(FrozenDictionary<TKey, TValue>))
        {
            return DictionaryBuilder<TKey, TValue>.FromDictionary(static values => values.ToFrozenDictionary(), copiesValues: true);
        }

        if (dictionaryType.IsInterface)
        {
            // Create a dictionary, and assign it to the member.
            return typeof(Dictionary<TKey, TValue>).IsAssignableTo(dictionaryType) ? DictionaryBuilder<TKey, TValue>.ForDictionary() : null;
        }

        if (dictionaryType.GetConstructor([]) is not null)
        {
            var create = CompiledAccessors.GetParameterlessConstructor(dictionaryType);
            if (DictionaryBuilder<TKey, TValue>.ForInterfaces(dictionaryType, create) is { } builder)
            {
                return builder;
            }

            // An Add(TKey, TValue) method, such as StringDictionary's.
            if (dictionaryType.GetMethod("Add", [typeof(TKey), typeof(TValue)]) is { } addMethod)
            {
                return DictionaryBuilder<TKey, TValue>.Direct(create, CompiledAccessors.GetAction<TKey, TValue?>(addMethod));
            }

            return null;
        }

        // A constructor that takes a dictionary, such as ReadOnlyDictionary<TKey, TValue>, which
        // keeps it.
        if (CollectionUtilities.GetConstructorTaking(dictionaryType, typeof(Dictionary<TKey, TValue>)) is { } constructor)
        {
            var construct = CompiledAccessors.GetConstructor1(constructor);
            return DictionaryBuilder<TKey, TValue>.FromDictionary(construct, copiesValues: false);
        }

        return null;
    }
}

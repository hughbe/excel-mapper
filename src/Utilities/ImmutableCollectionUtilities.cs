using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;

namespace ExcelMapper.Utilities;

internal static class ImmutableCollectionUtilities
{
    private const string ImmutableArrayTypeName = "System.Collections.Immutable.ImmutableArray";
    private const string ImmutableArrayGenericTypeName = "System.Collections.Immutable.ImmutableArray`1";

    private const string ImmutableListTypeName = "System.Collections.Immutable.ImmutableList";
    private const string ImmutableListGenericTypeName = "System.Collections.Immutable.ImmutableList`1";
    private const string ImmutableListGenericInterfaceTypeName = "System.Collections.Immutable.IImmutableList`1";

    private const string ImmutableStackTypeName = "System.Collections.Immutable.ImmutableStack";
    private const string ImmutableStackGenericTypeName = "System.Collections.Immutable.ImmutableStack`1";
    private const string ImmutableStackGenericInterfaceTypeName = "System.Collections.Immutable.IImmutableStack`1";

    private const string ImmutableQueueTypeName = "System.Collections.Immutable.ImmutableQueue";
    private const string ImmutableQueueGenericTypeName = "System.Collections.Immutable.ImmutableQueue`1";
    private const string ImmutableQueueGenericInterfaceTypeName = "System.Collections.Immutable.IImmutableQueue`1";

    private const string ImmutableSortedSetTypeName = "System.Collections.Immutable.ImmutableSortedSet";
    private const string ImmutableSortedSetGenericTypeName = "System.Collections.Immutable.ImmutableSortedSet`1";

    private const string ImmutableHashSetTypeName = "System.Collections.Immutable.ImmutableHashSet";
    private const string ImmutableHashSetGenericTypeName = "System.Collections.Immutable.ImmutableHashSet`1";
    private const string ImmutableSetGenericInterfaceTypeName = "System.Collections.Immutable.IImmutableSet`1";

    private const string ImmutableDictionaryTypeName = "System.Collections.Immutable.ImmutableDictionary";
    private const string ImmutableDictionaryGenericTypeName = "System.Collections.Immutable.ImmutableDictionary`2";
    private const string ImmutableDictionaryGenericInterfaceTypeName = "System.Collections.Immutable.IImmutableDictionary`2";

    private const string ImmutableSortedDictionaryTypeName = "System.Collections.Immutable.ImmutableSortedDictionary";
    private const string ImmutableSortedDictionaryGenericTypeName = "System.Collections.Immutable.ImmutableSortedDictionary`2";

    private static HashSet<string> ImmutableEnumerableTypeNames { get; } = new()
    {
        ImmutableArrayGenericTypeName,
        ImmutableListGenericTypeName,
        ImmutableListGenericInterfaceTypeName,
        ImmutableStackGenericTypeName,
        ImmutableStackGenericInterfaceTypeName,
        ImmutableQueueGenericTypeName,
        ImmutableQueueGenericInterfaceTypeName,
        ImmutableSortedSetGenericTypeName,
        ImmutableHashSetGenericTypeName,
        ImmutableSetGenericInterfaceTypeName
    };

    public static bool IsImmutableEnumerableType(this Type type) =>
        IsImmutableCollectionsType(type) && 
        ImmutableEnumerableTypeNames.Contains(type.GetGenericTypeDefinition().FullName);

    private static HashSet<string> ImmutableDictionaryTypeNames { get; } = new()
    {
        ImmutableDictionaryGenericTypeName,
        ImmutableDictionaryGenericInterfaceTypeName,
        ImmutableSortedDictionaryGenericTypeName
    };

    public static bool IsImmutableDictionaryType(this Type type) =>
        IsImmutableCollectionsType(type) && 
        ImmutableDictionaryTypeNames.Contains(type.GetGenericTypeDefinition().FullName);

    private static bool IsImmutableCollectionsType(Type type)
        => type.IsGenericType && type.Assembly.FullName.StartsWith("System.Collections.Immutable,", StringComparison.Ordinal);

    private static Dictionary<string, string> ImmutableEnumerableConstructingTypeMap { get; } = new()
    {
        [ImmutableArrayGenericTypeName] = ImmutableArrayTypeName,
        [ImmutableListGenericTypeName] = ImmutableListTypeName,
        [ImmutableListGenericInterfaceTypeName] = ImmutableListTypeName,
        [ImmutableStackGenericTypeName] = ImmutableStackTypeName,
        [ImmutableStackGenericInterfaceTypeName] = ImmutableStackTypeName,
        [ImmutableQueueGenericTypeName] = ImmutableQueueTypeName,
        [ImmutableQueueGenericInterfaceTypeName] = ImmutableQueueTypeName,
        [ImmutableSortedSetGenericTypeName] = ImmutableSortedSetTypeName,
        [ImmutableHashSetGenericTypeName] = ImmutableHashSetTypeName,
        [ImmutableSetGenericInterfaceTypeName] = ImmutableHashSetTypeName
    };


    private static Type GetImmutableEnumerableConstructingType(Type type)
    {
        Debug.Assert(type.IsImmutableEnumerableType());

        // Use the generic type definition of the immutable collection to determine
        // an appropriate constructing type, i.e. a type that we can invoke the
        // `CreateRange<T>` method on, which returns the desired immutable collection.
        Type underlyingType = type.GetGenericTypeDefinition();
        string fullName = underlyingType.FullName;

        Debug.Assert(ImmutableEnumerableConstructingTypeMap.ContainsKey(fullName), 
            $"Unknown type {fullName}");

        string constructingTypeName = ImmutableEnumerableConstructingTypeMap[fullName];

        // This won't be null because we verified the assembly is actually System.Collections.Immutable.
        return underlyingType.Assembly.GetType(constructingTypeName);
    }

    private static readonly Dictionary<string, string> ImmutableDictionaryConstructingTypeMap = new()
    {
        [ImmutableDictionaryGenericTypeName] = ImmutableDictionaryTypeName,
        [ImmutableDictionaryGenericInterfaceTypeName] = ImmutableDictionaryTypeName,
        [ImmutableSortedDictionaryGenericTypeName] = ImmutableSortedDictionaryTypeName
    };

    private static Type GetImmutableDictionaryConstructingType(Type type)
    {
        Debug.Assert(type.IsImmutableDictionaryType());

        // Use the generic type definition of the immutable collection to determine
        // an appropriate constructing type, i.e. a type that we can invoke the
        // `CreateRange<T>` method on, which returns the desired immutable collection.
        Type underlyingType = type.GetGenericTypeDefinition();
        string fullName = underlyingType.FullName;

        Debug.Assert(ImmutableDictionaryConstructingTypeMap.ContainsKey(fullName), 
            $"Unknown type {fullName}");

        string constructingTypeName = ImmutableDictionaryConstructingTypeMap[fullName];

        // This won't be null because we verified the assembly is actually System.Collections.Immutable.
        return underlyingType.Assembly.GetType(constructingTypeName);
    }

    public static MethodInfo GetImmutableEnumerableCreateRangeMethod(this Type type, Type elementType)
    {
        Type constructingType = GetImmutableEnumerableConstructingType(type);
        return constructingType.GetMethods()
            .First(m => m.Name == "CreateRange" && m.GetParameters().Length == 1)
            .MakeGenericMethod(elementType);
    }

    public static MethodInfo GetImmutableDictionaryCreateRangeMethod(this Type type, Type elementType)
    {
        Type constructingType = GetImmutableDictionaryConstructingType(type);
        return constructingType.GetMethods()
            .First(m => m.Name == "CreateRange" && m.GetParameters().Length == 1)
            .MakeGenericMethod(typeof(string), elementType);
    }
}

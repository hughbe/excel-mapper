using System;
using System.Diagnostics;
using System.Linq;
using System.Reflection;

namespace ExcelMapper.Utilities
{
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

        public static bool IsImmutableEnumerableType(this Type type)
        {
            if (!type.IsGenericType|| !type.Assembly.FullName.StartsWith("System.Collections.Immutable,", StringComparison.Ordinal))
            {
                return false;
            }

            switch (type.GetGenericTypeDefinition().FullName)
            {
                case ImmutableArrayGenericTypeName:
                case ImmutableListGenericTypeName:
                case ImmutableListGenericInterfaceTypeName:
                case ImmutableStackGenericTypeName:
                case ImmutableStackGenericInterfaceTypeName:
                case ImmutableQueueGenericTypeName:
                case ImmutableQueueGenericInterfaceTypeName:
                case ImmutableSortedSetGenericTypeName:
                case ImmutableHashSetGenericTypeName:
                case ImmutableSetGenericInterfaceTypeName:
                    return true;
                default:
                    return false;
            }
        }

        public static bool IsImmutableDictionaryType(this Type type)
        {
            if (!type.IsGenericType || !type.Assembly.FullName.StartsWith("System.Collections.Immutable,", StringComparison.Ordinal))
            {
                return false;
            }

            switch (type.GetGenericTypeDefinition().FullName)
            {
                case ImmutableDictionaryGenericTypeName:
                case ImmutableDictionaryGenericInterfaceTypeName:
                case ImmutableSortedDictionaryGenericTypeName:
                    return true;
                default:
                    return false;
            }
        }

        private static Type GetImmutableEnumerableConstructingType(Type type)
        {
            Debug.Assert(type.IsImmutableEnumerableType());

            // Use the generic type definition of the immutable collection to determine
            // an appropriate constructing type, i.e. a type that we can invoke the
            // `CreateRange<T>` method on, which returns the desired immutable collection.
            Type underlyingType = type.GetGenericTypeDefinition();
            string constructingTypeName;

            switch (underlyingType.FullName)
            {
                case ImmutableArrayGenericTypeName:
                    constructingTypeName = ImmutableArrayTypeName;
                    break;
                case ImmutableListGenericTypeName:
                case ImmutableListGenericInterfaceTypeName:
                    constructingTypeName = ImmutableListTypeName;
                    break;
                case ImmutableStackGenericTypeName:
                case ImmutableStackGenericInterfaceTypeName:
                    constructingTypeName = ImmutableStackTypeName;
                    break;
                case ImmutableQueueGenericTypeName:
                case ImmutableQueueGenericInterfaceTypeName:
                    constructingTypeName = ImmutableQueueTypeName;
                    break;
                case ImmutableSortedSetGenericTypeName:
                    constructingTypeName = ImmutableSortedSetTypeName;
                    break;
                default:
                    Debug.Assert(underlyingType.FullName == ImmutableHashSetGenericTypeName || underlyingType.FullName == ImmutableSetGenericInterfaceTypeName, $"Unknown type {underlyingType.FullName}");
                    constructingTypeName = ImmutableHashSetTypeName;
                    break;
            }

            // This won't be null because we verified the assembly is actually System.Collections.Immutable.
            return underlyingType.Assembly.GetType(constructingTypeName);
        }

        private static Type GetImmutableDictionaryConstructingType(Type type)
        {
            Debug.Assert(type.IsImmutableDictionaryType());

            // Use the generic type definition of the immutable collection to determine
            // an appropriate constructing type, i.e. a type that we can invoke the
            // `CreateRange<T>` method on, which returns the desired immutable collection.
            Type underlyingType = type.GetGenericTypeDefinition();
            string constructingTypeName;

            switch (underlyingType.FullName)
            {
                case ImmutableDictionaryGenericTypeName:
                case ImmutableDictionaryGenericInterfaceTypeName:
                    constructingTypeName = ImmutableDictionaryTypeName;
                    break;
                default:
                    Debug.Assert(underlyingType.FullName == ImmutableSortedDictionaryGenericTypeName, $"Unknown type {underlyingType.FullName}");
                    constructingTypeName = ImmutableSortedDictionaryTypeName;
                    break;
            }

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
}

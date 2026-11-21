using System.Reflection;
using System.Diagnostics.CodeAnalysis;
using System.Collections.Concurrent;

namespace ExcelMapper.Utilities;

internal static class CollectionUtilities
{
    /// <summary>
    /// Whether a collection type that's created by passing its items to its constructor (see
    /// <see cref="GetListConstructor{T}"/>) copies them, rather than keeping the list (as
    /// <see cref="System.Collections.ObjectModel.ReadOnlyCollection{T}"/> does), so the list can be
    /// reused for the next row. These are the built-in types that are created that way; other types,
    /// including the user's own, get a new list for each row.
    /// </summary>
    public static bool ConstructorCopies(Type type)
    {
        if (!type.IsGenericType)
        {
            return false;
        }

        var definition = type.GetGenericTypeDefinition();
        return definition == typeof(Queue<>)
            || definition == typeof(Stack<>)
            || definition == typeof(ConcurrentQueue<>)
            || definition == typeof(ConcurrentStack<>)
            || definition == typeof(ConcurrentBag<>);
    }

    /// <summary>
    /// Finds the public constructor of <paramref name="type"/> with one parameter that a
    /// <paramref name="argumentType"/> can be passed to, but not an <paramref name="excludedArgumentType"/>,
    /// preferring the most specific parameter type. <see cref="Type.GetConstructor(Type[])"/> isn't
    /// used because it also matches parameters that the given type can be assigned to, so a lookup
    /// for a constructor taking <c>ISet&lt;T&gt;</c> matches <c>Queue&lt;T&gt;(IEnumerable&lt;T&gt;)</c>.
    /// </summary>
    public static ConstructorInfo? GetConstructorTaking([DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicConstructors)] Type type, Type argumentType, Type? excludedArgumentType = null)
    {
        ConstructorInfo? result = null;
        Type? resultParameterType = null;
        foreach (var constructor in type.GetConstructors())
        {
            var parameters = constructor.GetParameters();
            if (parameters.Length != 1)
            {
                continue;
            }

            var parameterType = parameters[0].ParameterType;
            if (!parameterType.IsAssignableFrom(argumentType) ||
                (excludedArgumentType is not null && parameterType.IsAssignableFrom(excludedArgumentType)))
            {
                continue;
            }

            if (resultParameterType is null || resultParameterType.IsAssignableFrom(parameterType))
            {
                result = constructor;
                resultParameterType = parameterType;
            }
        }

        return result;
    }

    /// <summary>
    /// Finds a constructor that needs a set: one that takes a <see cref="HashSet{T}"/> but not a
    /// <see cref="List{T}"/>, such as <c>ReadOnlySet&lt;T&gt;(ISet&lt;T&gt;)</c>.
    /// </summary>
    public static ConstructorInfo? GetSetConstructor<T>([DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicConstructors)] Type type)
        => GetConstructorTaking(type, typeof(HashSet<T>), excludedArgumentType: typeof(List<T>));

    /// <summary>
    /// Finds a constructor that takes the items as a <see cref="List{T}"/>, for example through
    /// <c>IEnumerable&lt;T&gt;</c>, <c>IList&lt;T&gt;</c> or <see cref="System.Collections.ICollection"/>.
    /// </summary>
    public static ConstructorInfo? GetListConstructor<T>([DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicConstructors)] Type type)
        => GetConstructorTaking(type, typeof(List<T>));
}

using System.Diagnostics.CodeAnalysis;
using System.Collections.Concurrent;
using System.Reflection;

namespace ExcelMapper.AutoMapping;

/// <summary>
/// The attributes of members, read once per member. Mapping a member checks it for about 30
/// attribute types, and each <see cref="CustomAttributeExtensions.GetCustomAttribute{T}(MemberInfo)"/>
/// call reads the metadata again.
/// </summary>
/// <remarks>
/// The attribute instances are shared, so they must not be changed, and arrays they hold must be
/// copied before they're given to a map.
/// </remarks>
internal static class MemberAttributes
{
    private static readonly ConcurrentDictionary<MemberInfo, Attribute[]> s_attributes = new();

    // The same attributes as the CustomAttributeExtensions methods: inherited attributes included.
    private static Attribute[] GetAll(MemberInfo member)
        => s_attributes.GetOrAdd(member, static m => Attribute.GetCustomAttributes(m, inherit: true));

    /// <summary>
    /// Gets the attribute of type <typeparamref name="T"/>, like
    /// <see cref="CustomAttributeExtensions.GetCustomAttribute{T}(MemberInfo)"/>.
    /// </summary>
    /// <exception cref="AmbiguousMatchException">More than one attribute matches.</exception>
    public static T? Get<T>(MemberInfo member) where T : Attribute
    {
        T? result = null;
        foreach (var attribute in GetAll(member))
        {
            if (attribute is T match)
            {
                if (result is not null)
                {
                    throw new AmbiguousMatchException($"Multiple custom attributes of the same type '{typeof(T)}' found.");
                }

                result = match;
            }
        }

        return result;
    }

    /// <summary>
    /// Gets the attributes of type <typeparamref name="T"/>, in the order
    /// <see cref="CustomAttributeExtensions.GetCustomAttributes{T}(MemberInfo)"/> returns them.
    /// </summary>
    public static T[] GetMany<T>(MemberInfo member) where T : Attribute
    {
        var attributes = GetAll(member);
        var count = 0;
        foreach (var attribute in attributes)
        {
            if (attribute is T)
            {
                count++;
            }
        }

        if (count == 0)
        {
            return [];
        }

        var result = new T[count];
        var index = 0;
        foreach (var attribute in attributes)
        {
            if (attribute is T match)
            {
                result[index++] = match;
            }
        }

        return result;
    }

    /// <summary>
    /// Gets whether the member has an attribute of type <typeparamref name="T"/>.
    /// </summary>
    public static bool IsDefined<T>(MemberInfo member) where T : Attribute
    {
        foreach (var attribute in GetAll(member))
        {
            if (attribute is T)
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Creates the object an attribute gives the type of, such as the converter of an
    /// <see cref="ExcelConverterAttribute"/>, with the attribute's constructor arguments. If it
    /// can't be created, the <see cref="ExcelMappingException"/> names the type, the attribute and
    /// the member.
    /// </summary>
    public static object CreateInstance([DynamicallyAccessedMembers(TrimmingConstants.ActivatedType)] Type type, object?[]? arguments, string attributeName, MemberInfo member)
    {
        try
        {
            return Activator.CreateInstance(type, arguments)!;
        }
        catch (Exception exception) when (exception is MissingMethodException or AmbiguousMatchException or MemberAccessException or TargetInvocationException)
        {
            // A constructor that throws is reported as its own exception.
            var reason = exception is TargetInvocationException { InnerException: { } inner } ? inner : exception;
            throw new ExcelMappingException($"Can't create {TypeNames.Get(type)} given by {attributeName} on member \"{member.Name}\": {reason.Message}", reason);
        }
    }
}

using System.Diagnostics.CodeAnalysis;
using System.Collections.Concurrent;
using System.Globalization;
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

    /// <summary>
    /// Converts a value an attribute gives for a member, such as <see cref="ExcelDefaultValueAttribute.Value"/>,
    /// to the member's type when the class map is created. Attribute arguments are objects, and
    /// can't be <see cref="decimal"/>, so a number is converted to another number type if that
    /// doesn't change it: <c>[ExcelDefaultValue(0)]</c> works for a <see cref="double"/> or
    /// <see cref="decimal"/> member, and an integer for an enum member.
    /// </summary>
    /// <exception cref="ExcelMappingException">The value can't be the member's type.</exception>
    public static T ConvertValue<T>(object? value, string attributeName, MemberInfo member, string? key = null)
    {
        if (value is T typedValue)
        {
            return typedValue;
        }

        if (value is null)
        {
            if (default(T) is null)
            {
                return default!;
            }
        }
        else if (TryConvertNumber(value, Nullable.GetUnderlyingType(typeof(T)) ?? typeof(T), out var converted))
        {
            return (T)converted;
        }

        var givenFor = key is null ? string.Empty : $" for \"{key}\"";
        throw new ExcelMappingException($"The value {DescribeValue(value)} given by {attributeName}{givenFor} on member \"{member.Name}\" can't be converted to {TypeNames.Get(typeof(T))}.");
    }

    // A number converted to another number type, or an integer to an enum, if the value is the same.
    private static bool TryConvertNumber(object value, Type type, [NotNullWhen(true)] out object? converted)
    {
        converted = null;
        if (!IsNumber(value.GetType()))
        {
            return false;
        }

        if (type.IsEnum)
        {
            if (Type.GetTypeCode(value.GetType()) is TypeCode.Single or TypeCode.Double or TypeCode.Decimal)
            {
                return false;
            }

            converted = Enum.ToObject(type, value);
            return true;
        }

        if (!IsNumber(type))
        {
            return false;
        }

        try
        {
            var result = Convert.ChangeType(value, type, CultureInfo.InvariantCulture);

            // Rounding (1.5 to an int) changes the value, so it isn't a conversion.
            if (!Equals(Convert.ChangeType(result, value.GetType(), CultureInfo.InvariantCulture), value))
            {
                return false;
            }

            converted = result;
            return true;
        }
        catch (OverflowException)
        {
            return false;
        }
    }

    private static bool IsNumber(Type type) => Type.GetTypeCode(type) is >= TypeCode.SByte and <= TypeCode.Decimal;

    private static string DescribeValue(object? value) => value switch
    {
        null => "null",
        string text => $"\"{text}\"",
        _ => Convert.ToString(value, CultureInfo.InvariantCulture) ?? value.GetType().Name,
    };
}

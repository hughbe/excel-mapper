using System.Reflection;

namespace ExcelMapper.Utilities;

/// <summary>
/// Helpers for creating objects through a constructor whose parameters are populated from mapped
/// members, for example positional records or classes with get-only properties.
/// </summary>
internal static class ConstructorUtilities
{
    /// <summary>
    /// Gets whether the type can be created without passing any constructor arguments.
    /// </summary>
    public static bool HasParameterlessConstructor(Type type)
        => type.IsValueType || type.GetConstructor(Type.EmptyTypes) is not null;

    /// <summary>
    /// Gets the type of a property or field.
    /// </summary>
    public static Type GetMemberType(MemberInfo member) => member switch
    {
        PropertyInfo property => property.PropertyType,
        FieldInfo field => field.FieldType,
        _ => throw new ArgumentException($"Member \"{member.Name}\" is not a field or property.", nameof(member))
    };

    /// <summary>
    /// Gets whether a public constructor of the type has a parameter that a member with the given
    /// name and type could be passed to.
    /// </summary>
    public static bool IsConstructorParameter(Type type, string memberName, Type memberType)
    {
        foreach (var constructor in type.GetConstructors())
        {
            foreach (var parameter in constructor.GetParameters())
            {
                if (ParameterMatches(parameter, memberName, memberType))
                {
                    return true;
                }
            }
        }

        return false;
    }

    private static bool ParameterMatches(ParameterInfo parameter, string memberName, Type memberType)
        => string.Equals(parameter.Name, memberName, StringComparison.OrdinalIgnoreCase)
            && parameter.ParameterType.IsAssignableFrom(memberType);

    /// <summary>
    /// Finds the public constructor (with at least one parameter) that best matches the given members.
    /// Every parameter must either match a member by name (case-insensitively) and type, or have a
    /// default value. The constructor matching the most members is chosen.
    /// </summary>
    /// <param name="type">The type to construct.</param>
    /// <param name="members">The members whose values are available.</param>
    /// <param name="parameterMemberIndices">For each parameter of the returned constructor, the index
    /// of the matching member in <paramref name="members"/>, or -1 if the default value is used.</param>
    /// <returns>The best matching constructor, or null if there is none.</returns>
    public static ConstructorInfo? FindConstructor(Type type, IReadOnlyList<MemberInfo> members, out int[] parameterMemberIndices)
    {
        ConstructorInfo? bestConstructor = null;
        int[] bestIndices = [];
        int bestMatched = -1;
        int bestParameterCount = 0;

        foreach (var constructor in type.GetConstructors())
        {
            var parameters = constructor.GetParameters();
            if (parameters.Length == 0)
            {
                continue;
            }

            var indices = new int[parameters.Length];
            var matched = 0;
            var usable = true;
            for (int i = 0; i < parameters.Length; i++)
            {
                indices[i] = -1;
                for (int j = 0; j < members.Count; j++)
                {
                    if (ParameterMatches(parameters[i], members[j].Name, GetMemberType(members[j])))
                    {
                        indices[i] = j;
                        matched++;
                        break;
                    }
                }

                if (indices[i] == -1 && !parameters[i].HasDefaultValue)
                {
                    usable = false;
                    break;
                }
            }

            if (!usable)
            {
                continue;
            }

            // Prefer the constructor that uses the most mapped members, then the one with the
            // fewest parameters left at their default values.
            if (matched > bestMatched || (matched == bestMatched && parameters.Length < bestParameterCount))
            {
                bestConstructor = constructor;
                bestIndices = indices;
                bestMatched = matched;
                bestParameterCount = parameters.Length;
            }
        }

        parameterMemberIndices = bestIndices;
        return bestConstructor;
    }

    /// <summary>
    /// Gets the public instance properties and fields of a type whose values could be passed to a
    /// constructor.
    /// </summary>
    public static List<MemberInfo> GetReadableMembers(Type type)
    {
        var members = new List<MemberInfo>();
        foreach (var property in type.GetProperties(BindingFlags.Public | BindingFlags.Instance))
        {
            if (property.GetMethod is { IsPublic: true } && property.GetIndexParameters().Length == 0)
            {
                members.Add(property);
            }
        }

        foreach (var field in type.GetFields(BindingFlags.Public | BindingFlags.Instance))
        {
            members.Add(field);
        }

        return members;
    }
}

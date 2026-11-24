using System.Diagnostics.CodeAnalysis;
using System.Reflection;

namespace ExcelMapper.Maps;

/// <summary>
/// A member of a class map (<see cref="ExcelClassMap.Members"/>): the property or field, and the map
/// that reads its value.
/// </summary>
internal sealed class ClassMapMember
{
    /// <summary>
    /// Creates a member of a class map.
    /// </summary>
    /// <param name="member">The property or field of the class map's type.</param>
    /// <param name="map">The map that reads the member's value.</param>
    /// <param name="mappedExplicitly">Whether the member was mapped with <c>Map(...)</c>, rather
    /// than by auto-mapping.</param>
    /// <exception cref="ArgumentException"><paramref name="member"/> is a read-only property that
    /// doesn't match a constructor parameter.</exception>
    [UnconditionalSuppressMessage("Trimming", "IL2072", Justification =
        "Only read-only properties check the constructors of their type, which is the type of the class map " +
        "they're added to, and ExcelClassMap keeps that type's public constructors. If they were removed anyway, " +
        "this throws an ArgumentException rather than mapping incorrectly.")]
    public ClassMapMember(MemberInfo member, IMap map, bool mappedExplicitly = false)
    {
        // Members come from expressions on the mapped object and from its public instance
        // properties and fields, so they're never indexers or static.
        Member = member;
        Map = map;
        MappedExplicitly = mappedExplicitly;

        if (member is PropertyInfo { CanWrite: false } property)
        {
            // A read-only property can only be mapped if its value can be passed to a
            // constructor parameter with the same name (for example a positional record).
            var type = property.ReflectedType ?? property.DeclaringType!;
            if (property.GetMethod is null || property.GetMethod.IsStatic ||
                !ConstructorUtilities.IsConstructorParameter(type, property.Name, property.PropertyType))
            {
                throw new ArgumentException($"Property \"{member.Name}\" is read-only and does not match a constructor parameter.", nameof(member));
            }

            IsReadOnly = true;
        }
    }

    /// <summary>
    /// Gets the member that is mapped.
    /// </summary>
    public MemberInfo Member { get; }

    /// <summary>
    /// Gets whether the member is a read-only property that is set through a constructor parameter.
    /// </summary>
    public bool IsReadOnly { get; }

    /// <summary>
    /// Gets the map used to map the member.
    /// </summary>
    public IMap Map { get; }

    /// <summary>
    /// Gets whether the member was mapped with <c>Map(...)</c>, so that mapping it again configures
    /// the same map, rather than by auto-mapping, which <c>Map(...)</c> replaces.
    /// </summary>
    public bool MappedExplicitly { get; }
}

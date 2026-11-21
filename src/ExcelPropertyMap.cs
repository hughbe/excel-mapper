using System.Diagnostics.CodeAnalysis;
using System.Reflection;

namespace ExcelMapper;

/// <summary>
/// A delegate that sets the value of a member (property or field) on an object instance.
/// </summary>
/// <param name="instance">The object instance whose member value should be set.</param>
/// <param name="value">The value to set on the member.</param>
internal delegate void MemberSetValueDelegate(object instance, object value);

/// <summary>
/// Maps a member to an Excel column.
/// </summary>
internal class ExcelPropertyMap
{
    /// <summary>
    /// Constructs a property map for a member and map.
    /// </summary>
    /// <param name="member">The property or field of the class map's type.</param>
    /// <param name="map">The map that reads the member's value.</param>
    /// <exception cref="ArgumentException"><paramref name="member"/> is a read-only property that
    /// doesn't match a constructor parameter.</exception>
    [UnconditionalSuppressMessage("Trimming", "IL2072", Justification =
        "Only read-only properties check the constructors of their type, which is the type of the class map " +
        "they're added to, and ExcelClassMap keeps that type's public constructors. If they were removed anyway, " +
        "this throws an ArgumentException rather than mapping incorrectly.")]
    public ExcelPropertyMap(MemberInfo member, IMap map)
    {
        // Members come from expressions on the mapped object and from its public instance
        // properties and fields, so they're never indexers or static.
        Member = member;
        Map = map;

        if (member is PropertyInfo property)
        {
            if (property.CanWrite)
            {
                SetValueFactory = CompiledAccessors.GetSetter(property);
            }
            else
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
                SetValueFactory = (_, _) => throw new InvalidOperationException($"Property \"{member.Name}\" is read-only and can only be set through a constructor parameter.");
            }
        }
        else
        {
            SetValueFactory = CompiledAccessors.GetSetter((FieldInfo)member);
        }
    }

    /// <summary>
    /// Gets the member that is mapped.
    /// </summary>
    public MemberInfo Member { get; }

    /// <summary>
    /// Gets the delegate used to set the value of the member.
    /// </summary>
    public MemberSetValueDelegate SetValueFactory { get; }

    /// <summary>
    /// Gets whether the member is a read-only property that is set through a constructor parameter
    /// rather than <see cref="SetValueFactory"/>.
    /// </summary>
    public bool IsReadOnly { get; }

    /// <summary>
    /// Gets the map used to map the member.
    /// </summary>
    public IMap Map { get; }
}

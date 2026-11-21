using System.Diagnostics.CodeAnalysis;
using System.Reflection;

namespace ExcelMapper;

/// <summary>
/// A delegate that sets the value of a member (property or field) on an object instance.
/// </summary>
/// <param name="instance">The object instance whose member value should be set.</param>
/// <param name="value">The value to set on the member.</param>
public delegate void MemberSetValueDelegate(object instance, object value);

/// <summary>
/// Maps a member to an Excel column.
/// </summary>
public class ExcelPropertyMap
{
    /// <summary>
    /// Constructs a property map for a member and map.
    /// </summary>
    /// <param name="member"></param>
    /// <param name="map"></param>
    /// <exception cref="ArgumentException"></exception>
    [UnconditionalSuppressMessage("Trimming", "IL2072", Justification =
        "Only read-only properties check the constructors of their type, which is the type of the class map " +
        "they're added to, and ExcelClassMap keeps that type's public constructors. If they were removed anyway, " +
        "this throws an ArgumentException rather than mapping incorrectly.")]
    public ExcelPropertyMap(MemberInfo member, IMap map)
    {
        ThrowHelpers.ThrowIfNull(member, nameof(member));
        ThrowHelpers.ThrowIfNull(map, nameof(map));

        Member = member;
        Map = map;

        if (member is PropertyInfo property)
        {
            // Property must not be an indexer.
            if (property.GetIndexParameters().Length > 0)
            {
                throw new ArgumentException($"Property \"{member.Name}\" is an indexer and cannot be mapped.", nameof(member));
            }

            if (property.CanWrite)
            {
                // Property must be an instance property.
                if (property.SetMethod!.IsStatic)
                {
                    throw new ArgumentException($"Property \"{member.Name}\" cannot be static.", nameof(member));
                }

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
        else if (member is FieldInfo field)
        {
            SetValueFactory = CompiledAccessors.GetSetter(field);
        }
        else
        {
            throw new ArgumentException($"Member \"{member.Name}\" is not a field or property.", nameof(member));
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

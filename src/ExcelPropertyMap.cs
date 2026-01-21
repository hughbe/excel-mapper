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
    public ExcelPropertyMap(MemberInfo member, IMap map)
    {
        ThrowHelpers.ThrowIfNull(member, nameof(member));
        ThrowHelpers.ThrowIfNull(map, nameof(map));

        Member = member;
        Map = map;

        if (member is PropertyInfo property)
        {
            // Property must have a setter.
            if (!property.CanWrite)
            {
                throw new ArgumentException($"Property \"{member.Name}\" is read-only.", nameof(member));
            }
            // Property must be an instance property.
            if (property.SetMethod!.IsStatic)
            {
                throw new ArgumentException($"Property \"{member.Name}\" cannot be static.", nameof(member));
            }
            // Property must not be an indexer.
            if (property.GetIndexParameters().Length > 0)
            {
                throw new ArgumentException($"Property \"{member.Name}\" is an indexer and cannot be mapped.", nameof(member));
            }

            SetValueFactory = property.SetValue;
        }
        else if (member is FieldInfo field)
        {
            SetValueFactory = field.SetValue;
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
    /// Gets the map used to map the member.
    /// </summary>
    public IMap Map { get; }
}

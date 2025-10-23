using System.Reflection;

namespace ExcelMapper;

/// <summary>
/// Maps a member to an Excel column of type T.
/// </summary>
/// <typeparam name="T">The type of the member.</typeparam>
public class ExcelPropertyMap<T> : ExcelPropertyMap
{
    /// <summary>
    /// Constructs a property map for a member and map.
    /// </summary>
    /// <param name="member">The member to map.</param>
    /// <param name="map">The map to use.</param>
    public ExcelPropertyMap(MemberInfo member, IMap map) : base(member, map)
    {
    }
}

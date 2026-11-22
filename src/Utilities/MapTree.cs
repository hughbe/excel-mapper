using System.Reflection;

namespace ExcelMapper.Utilities;

/// <summary>
/// Walks a tree of maps: class maps and members mapped one by one (<see cref="IParentMap"/>)
/// down to the maps that read cells.
/// </summary>
internal static class MapTree
{
    /// <summary>
    /// Calls <paramref name="visit"/> for each map under <paramref name="map"/> that reads cells
    /// (single values, collections and dictionaries), with the path of the member it's for.
    /// </summary>
    public static void VisitCellMaps(IMap map, string? path, Action<IToOneMap, string?> visit)
    {
        switch (map)
        {
            case IParentMap parent:
                foreach (var child in parent.Children)
                {
                    VisitCellMaps(child.Value, parent.GetChildPath(path, child.Key), visit);
                }

                break;

            case IToOneMap cellMap:
                visit(cellMap, path);
                break;
        }
    }

    /// <summary>
    /// Finds the class map, under <paramref name="parent"/>, that has a member mapped by
    /// <paramref name="map"/>.
    /// </summary>
    public static ExcelClassMap? FindOwner(IMap parent, IMap map, out MemberInfo? member)
    {
        if (parent is IParentMap parentMap)
        {
            foreach (var child in parentMap.Children)
            {
                if (ReferenceEquals(child.Value, map) && parent is ExcelClassMap classMap && child.Key is MemberInfo childMember)
                {
                    member = childMember;
                    return classMap;
                }

                if (FindOwner(child.Value, map, out member) is { } owner)
                {
                    return owner;
                }
            }
        }

        member = null;
        return null;
    }
}

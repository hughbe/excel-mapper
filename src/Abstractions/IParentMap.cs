namespace ExcelMapper.Abstractions;

/// <summary>
/// A map made of other maps: a class map's members, or the elements or values of a member mapped
/// one by one, such as <c>Map(p => p.Values[0])</c>. <see cref="MapTree"/> walks them.
/// </summary>
internal interface IParentMap : IMap
{
    /// <summary>
    /// Gets the child maps and their keys: members for a class map (a class map read as a single
    /// value has a child with no key), and indices or dictionary keys for the others.
    /// </summary>
    IEnumerable<KeyValuePair<object?, IMap>> Children { get; }

    /// <summary>
    /// Gets the child map with the given key, or null if there isn't one.
    /// </summary>
    IMap? GetChild(object key);

    /// <summary>
    /// Gets the child map with the given key if it was mapped with <c>Map(...)</c>, or null if
    /// there isn't one or auto-mapping created it: mapping a member again configures the map it
    /// already has (see <c>ExpressionAutoMapper</c>), but replaces an auto-mapped one.
    /// </summary>
    IMap? GetMappedChild(object key);

    /// <summary>
    /// Sets the child map with the given key, replacing any it has.
    /// </summary>
    void SetChild(object key, IMap map);

    /// <summary>
    /// Gets the path of a child, for messages: the member's name (<c>Address.Street</c>), or the
    /// index or key (<c>Values[0]</c>, <c>Grid[0, 1]</c>, <c>Scores["Maths"]</c>).
    /// </summary>
    string? GetChildPath(string? path, object? key);
}

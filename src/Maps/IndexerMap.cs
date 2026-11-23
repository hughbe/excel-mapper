using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelDataReader;

namespace ExcelMapper.Maps;

/// <summary>
/// A collection, dictionary or array whose elements are mapped one by one, such as
/// <c>Map(p => p.Values[0])</c>, <c>Map(p => p.Scores["Maths"])</c> or <c>Map(p => p.Grid[0, 1])</c>:
/// a map for each index or key (<see cref="IParentMap"/>), whose values are put together when a
/// row is read.
/// </summary>
/// <typeparam name="TKey">The type of the index or key.</typeparam>
internal abstract class IndexerMap<TKey>(IEqualityComparer<TKey>? comparer = null) : IParentMap where TKey : notnull
{
    private bool _frozen;

    /// <summary>
    /// Gets the map of each element, by its index or key.
    /// </summary>
    public Dictionary<TKey, IMap> Values { get; } = new(comparer);

    IEnumerable<KeyValuePair<object?, IMap>> IParentMap.Children
    {
        get
        {
            foreach (var value in Values)
            {
                yield return new(value.Key, value.Value);
            }
        }
    }

    IMap? IParentMap.GetChild(object key) => Values.TryGetValue((TKey)key, out var map) ? map : null;

    // Elements mapped one by one are only ever mapped with Map(...).
    IMap? IParentMap.GetMappedChild(object key) => ((IParentMap)this).GetChild(key);

    void IParentMap.SetChild(object key, IMap map)
    {
        ThrowHelpers.ThrowIfFrozen(_frozen);
        Values[(TKey)key] = map;
    }

    string? IParentMap.GetChildPath(string? path, object? key) => $"{path}[{FormatKey((TKey)key!)}]";

    /// <summary>
    /// Formats an index or key for messages, as in C#: <c>0</c>, <c>"Maths"</c> or <c>0, 1</c>.
    /// </summary>
    private protected abstract string FormatKey(TKey key);

    void IMap.Freeze()
    {
        if (_frozen)
        {
            return;
        }

        _frozen = true;
        foreach (var map in Values.Values)
        {
            map.Freeze();
        }
    }

    /// <inheritdoc />
    public abstract bool TryGetValue(SheetView sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, ErrorTarget? errors, [NotNullWhen(true)] out object? value);
}

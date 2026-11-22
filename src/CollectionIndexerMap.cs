using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Factories;

namespace ExcelMapper;

/// <summary>
/// A map that reads multiple cells of an excel sheet and maps the values of the cells to a
/// list property or field by their indices.
/// </summary>
internal class CollectionIndexerMap<TValue> : IParentMap
{
    internal CollectionIndexerMap(CollectionBuilder<TValue> builder)
    {
        Builder = builder;
    }

    internal CollectionBuilder<TValue> Builder { get; }

    public Dictionary<int, IMap> Values { get; } = [];

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

    IMap? IParentMap.GetChild(object key) => Values.TryGetValue((int)key, out var map) ? map : null;

    void IParentMap.SetChild(object key, IMap map)
    {
        ThrowHelpers.ThrowIfFrozen(_frozen);
        Values[(int)key] = map;
    }

    private bool _frozen;

    /// <inheritdoc />
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

    string? IParentMap.GetChildPath(string? path, object? key) => $"{path}[{key}]";

    private int GetLength()
    {
        // The collection must be large enough for the largest index.
        var length = Values.Count;
        foreach (var index in Values.Keys)
        {
            length = Math.Max(length, index + 1);
        }

        return length;
    }

    public bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, List<ExcelMappingException>? errors, [NotNullWhen(true)] out object? value)
    {
        ThrowHelpers.ThrowIfNull(sheet, nameof(sheet));

        var items = Builder.Begin(GetLength());
        foreach (var map in Values)
        {
            if (map.Value.TryGetValue(sheet, rowIndex, reader, member, errors, out var elementValue))
            {
                Builder.Set(items, map.Key, (TValue)elementValue);
            }
        }

        value = Builder.End(items);
        return true;
    }
}

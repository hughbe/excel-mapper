using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Builders;

namespace ExcelMapper;

/// <summary>
/// A map that reads multiple cells of an excel sheet and maps the values of the cells to
/// a dictionary property or field by their keys.
/// </summary>
internal class DictionaryIndexerMap<TKey, TValue> : IParentMap where TKey : notnull
{
    internal DictionaryIndexerMap(DictionaryBuilder<TKey, TValue> builder)
    {
        Builder = builder;
    }

    internal DictionaryBuilder<TKey, TValue> Builder { get; }

    public Dictionary<object, IMap> Values { get; } = [];

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

    IMap? IParentMap.GetChild(object key) => Values.TryGetValue(key, out var map) ? map : null;

    void IParentMap.SetChild(object key, IMap map)
    {
        ThrowHelpers.ThrowIfFrozen(_frozen);
        Values[key] = map;
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

    string? IParentMap.GetChildPath(string? path, object? key) => $"{path}[\"{key}\"]";

    public bool TryGetValue(SheetView sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, List<ExcelMappingException>? errors, [NotNullWhen(true)] out object? value)
    {
        ThrowHelpers.ThrowIfNull(sheet, nameof(sheet));

        var items = Builder.Begin(Values.Count);
        foreach (var map in Values)
        {
            if (map.Value.TryGetValue(sheet, rowIndex, reader, member, errors, out var elementValue))
            {
                Builder.Add(items, (TKey)map.Key, (TValue?)elementValue);
            }
        }

        value = Builder.End(items);
        return true;
    }
}

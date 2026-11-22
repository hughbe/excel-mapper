using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Factories;

namespace ExcelMapper;

/// <summary>
/// A map that reads multiple cells of an excel sheet and maps the values of the cells to
/// a dictionary property or field by their keys.
/// </summary>
internal class DictionaryIndexerMap<TKey, TValue> : IDictionaryIndexerMap where TKey : notnull
{
    internal DictionaryIndexerMap(DictionaryBuilder<TKey, TValue> builder)
    {
        Builder = builder;
    }

    internal DictionaryBuilder<TKey, TValue> Builder { get; }

    public Dictionary<object, IMap> Values { get; } = [];

    public bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, List<ExcelMappingException>? errors, [NotNullWhen(true)] out object? value)
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

using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Builders;

namespace ExcelMapper.Maps;

/// <summary>
/// A map that reads multiple cells of an excel sheet and maps the values of the cells to
/// a dictionary property or field by their keys.
/// </summary>
internal sealed class DictionaryIndexerMap<TKey, TValue>(DictionaryBuilder<TKey, TValue> builder) : IndexerMap<object> where TKey : notnull
{

    internal DictionaryBuilder<TKey, TValue> Builder { get; } = builder;

    private protected override string FormatKey(object key) => $"\"{key}\"";

    public override bool TryGetValue(SheetView sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, ErrorTarget? errors, [NotNullWhen(true)] out object? value)
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

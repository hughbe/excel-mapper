using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Builders;

namespace ExcelMapper;

/// <summary>
/// A map that reads multiple cells of an excel sheet and maps the values of the cells to a
/// list property or field by their indices.
/// </summary>
internal sealed class CollectionIndexerMap<TValue>(CollectionBuilder<TValue> builder) : IndexerMap<int>
{

    internal CollectionBuilder<TValue> Builder { get; } = builder;

    private protected override string FormatKey(int key) => key.ToString(System.Globalization.CultureInfo.InvariantCulture);

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

    public override bool TryGetValue(SheetView sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, List<ExcelMappingException>? errors, [NotNullWhen(true)] out object? value)
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

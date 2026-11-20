using System.Runtime.CompilerServices;

namespace ExcelMapper.Utilities;

/// <summary>
/// Caches a value per sheet without keeping the sheet (and its importer and reader) alive.
/// </summary>
/// <remarks>
/// Maps are often registered in many importers, for example a single class map instance shared
/// by a long-running service. A strongly-keyed dictionary would keep every sheet ever read alive.
/// </remarks>
internal sealed class PerSheetCache<T> where T : class
{
    private sealed class Entry(T? value)
    {
        public T? Value { get; } = value;
    }

    private ConditionalWeakTable<ExcelSheet, Entry> _table = new();

    public T? GetOrAdd(ExcelSheet sheet, Func<ExcelSheet, T?> factory)
    {
        var table = _table;
        if (table.TryGetValue(sheet, out var entry))
        {
            return entry.Value;
        }

        return table.GetValue(sheet, s => new Entry(factory(s))).Value;
    }

    /// <summary>
    /// Discards every cached value, for example when the factory used to create them changes.
    /// </summary>
    public void Clear() => _table = new ConditionalWeakTable<ExcelSheet, Entry>();
}

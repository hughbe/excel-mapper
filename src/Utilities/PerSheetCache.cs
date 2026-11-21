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
        return table.TryGetValue(sheet, out var entry) ? entry.Value : Add(table, sheet, factory);
    }

    /// <summary>
    /// Gets the value for a sheet, creating it with <paramref name="factory"/> and <paramref name="state"/>
    /// if needed. Pass a static lambda so that looking up a cached value doesn't allocate.
    /// </summary>
    public T? GetOrAdd<TState>(ExcelSheet sheet, TState state, Func<ExcelSheet, TState, T?> factory)
    {
        var table = _table;
        return table.TryGetValue(sheet, out var entry) ? entry.Value : Add(table, sheet, state, factory);
    }

    // Separate methods so the closures are only allocated when a value is created. A lambda that
    // captures a parameter allocates its closure when the method is entered, even if it isn't used.
    private static T? Add(ConditionalWeakTable<ExcelSheet, Entry> table, ExcelSheet sheet, Func<ExcelSheet, T?> factory)
        => table.GetValue(sheet, s => new Entry(factory(s))).Value;

    private static T? Add<TState>(ConditionalWeakTable<ExcelSheet, Entry> table, ExcelSheet sheet, TState state, Func<ExcelSheet, TState, T?> factory)
        => table.GetValue(sheet, s => new Entry(factory(s, state))).Value;

    public void Clear() => _table = new ConditionalWeakTable<ExcelSheet, Entry>();
}

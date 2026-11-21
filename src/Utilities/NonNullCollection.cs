using System.Collections.ObjectModel;

namespace ExcelMapper.Utilities;

internal class NonNullCollection<T> : Collection<T> where T : class
{
    /// <summary>
    /// Incremented when the collection changes, so that anything built from it can be rebuilt.
    /// </summary>
    public int Version { get; private set; }

    protected override void RemoveItem(int index)
    {
        base.RemoveItem(index);
        Version++;
    }

    protected override void ClearItems()
    {
        base.ClearItems();
        Version++;
    }

    protected override void InsertItem(int index, T item)
    {
        ThrowHelpers.ThrowIfNull(item, nameof(item));
        base.InsertItem(index, item);
        Version++;
    }

    protected override void SetItem(int index, T item)
    {
        ThrowHelpers.ThrowIfNull(item, nameof(item));
        base.SetItem(index, item);
        Version++;
    }
}
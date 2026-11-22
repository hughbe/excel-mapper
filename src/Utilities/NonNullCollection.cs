using System.Collections.ObjectModel;

namespace ExcelMapper.Utilities;

/// <summary>
/// A collection that doesn't allow null items, and can't be changed once its map is frozen (see
/// <see cref="IMap.Freeze"/>).
/// </summary>
internal class NonNullCollection<T> : Collection<T> where T : class
{
    private bool _frozen;

    /// <summary>
    /// Stops the collection changing: its map has been used to read rows.
    /// </summary>
    public void Freeze() => _frozen = true;

    protected override void RemoveItem(int index)
    {
        ThrowHelpers.ThrowIfFrozen(_frozen);
        base.RemoveItem(index);
    }

    protected override void ClearItems()
    {
        ThrowHelpers.ThrowIfFrozen(_frozen);
        base.ClearItems();
    }

    protected override void InsertItem(int index, T item)
    {
        ThrowHelpers.ThrowIfFrozen(_frozen);
        ThrowHelpers.ThrowIfNull(item, nameof(item));
        base.InsertItem(index, item);
    }

    protected override void SetItem(int index, T item)
    {
        ThrowHelpers.ThrowIfFrozen(_frozen);
        ThrowHelpers.ThrowIfNull(item, nameof(item));
        base.SetItem(index, item);
    }
}

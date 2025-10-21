using System;
using System.Collections.ObjectModel;

namespace ExcelMapper.Utilities;

internal class NonNullCollection<T> : Collection<T> where T : class
{
    protected override void InsertItem(int index, T item)
    {
        ArgumentNullException.ThrowIfNull(item);
        base.InsertItem(index, item);
    }

    protected override void SetItem(int index, T item)
    {
        ArgumentNullException.ThrowIfNull(item);
        base.SetItem(index, item);
    }
}
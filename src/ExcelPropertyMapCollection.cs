using System;
using System.Collections.ObjectModel;

namespace ExcelMapper;

/// <summary>
/// A collection of property maps used by a class map.
/// </summary>
public class ExcelPropertyMapCollection : Collection<ExcelPropertyMap>
{
    protected override void InsertItem(int index, ExcelPropertyMap item)
    {
        ArgumentNullException.ThrowIfNull(item);

        base.InsertItem(index, item);
    }

    protected override void SetItem(int index, ExcelPropertyMap item)
    {
        ArgumentNullException.ThrowIfNull(item);

        base.SetItem(index, item);
    }
}

using System.Collections.ObjectModel;
namespace ExcelMapper.Abstractions;

/// <summary>
/// A collection of property maps used by a class map.
/// </summary>
public class CellValueMapperCollection : Collection<ICellValueMapper>
{
    protected override void InsertItem(int index, ICellValueMapper item)
    {
        if (item == null)
        {
            throw new ArgumentNullException(nameof(item));
        }

        base.InsertItem(index, item);
    }

    protected override void SetItem(int index, ICellValueMapper item)
    {
        if (item == null)
        {
            throw new ArgumentNullException(nameof(item));
        }

        base.SetItem(index, item);
    }
}

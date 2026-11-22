using System.Linq;
using System.Reflection;

namespace ExcelMapper.Readers;

/// <summary>
/// The columns an <see cref="IColumnSelector"/> chooses. In a region, columns outside it are ignored.
/// </summary>
internal sealed class SelectorColumns(IColumnSelector selector) : Columns
{
    public IColumnSelector Selector { get; } = selector;

    public override int FindColumn(SheetView sheet)
    {
        var columns = Select(sheet);
        return columns.Length == 0 ? -1 : columns[0];
    }

    public override int[]? FindColumns(SheetView sheet)
    {
        var columns = Select(sheet);
        return columns.Length == 0 ? null : columns;
    }

    // The chosen columns that are in the sheet (or region).
    private int[] Select(SheetView sheet)
    {
        var selected = Selector.SelectColumns(new ColumnSelectorContext(sheet)) ?? throw new InvalidOperationException($"The column selector \"{Selector.GetType()}\" returned null.");
        var end = Math.Min(sheet.ColumnEnd, sheet.NumberOfColumns);
        var columns = new List<int>(selected.Count);
        for (int i = 0; i < selected.Count; i++)
        {
            if (selected[i] >= sheet.ColumnOffset && selected[i] < end)
            {
                columns.Add(selected[i]);
            }
        }

        return [.. columns];
    }

    public override ExcelMappingException CreateMissingException(SheetView sheet, int rowIndex, MemberInfo? member)
        => Missing(sheet, rowIndex, member, Selector.Describe(new ColumnSelectorContext(sheet)));

    public override string Describe(SheetView sheet) => Selector.Describe(new ColumnSelectorContext(sheet));
}

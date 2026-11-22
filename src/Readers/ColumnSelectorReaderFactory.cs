namespace ExcelMapper.Readers;

/// <summary>
/// The reader factory for an <see cref="IColumnSelector"/>: a member with a single value reads the
/// first column chosen, and a collection or dictionary reads all of them.
/// </summary>
internal sealed class ColumnSelectorReaderFactory(IColumnSelector selector) : ICellReaderFactory, ICellsReaderFactory
{
    public IColumnSelector Selector { get; } = selector;

    public ICellReader? GetCellReader(ExcelSheet sheet)
    {
        ThrowHelpers.ThrowIfNull(sheet, nameof(sheet));
        var columns = SelectColumns(sheet);
        return columns.Length == 0 ? null : new ColumnIndexReader(columns[0]);
    }

    public ICellsReader? GetCellsReader(ExcelSheet sheet)
    {
        ThrowHelpers.ThrowIfNull(sheet, nameof(sheet));
        var columns = SelectColumns(sheet);
        return columns.Length == 0 ? null : new ColumnIndicesReader(columns);
    }

    // The chosen columns that are in the sheet (or region).
    private int[] SelectColumns(ExcelSheet sheet)
    {
        var selected = Selector.SelectColumns(new ColumnSelectorContext(sheet)) ?? throw new InvalidOperationException($"The column selector \"{Selector.GetType()}\" returned null.");
        var end = Math.Min(sheet.ColumnEnd, sheet.NumberOfColumns);
        var count = 0;
        for (int i = 0; i < selected.Count; i++)
        {
            if (selected[i] >= sheet.ColumnOffset && selected[i] < end)
            {
                count++;
            }
        }

        var columns = new int[count];
        count = 0;
        for (int i = 0; i < selected.Count; i++)
        {
            if (selected[i] >= sheet.ColumnOffset && selected[i] < end)
            {
                columns[count++] = selected[i];
            }
        }

        return columns;
    }
}

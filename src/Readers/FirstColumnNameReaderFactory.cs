namespace ExcelMapper.Readers;

/// <summary>
/// Reads a member with several <see cref="ExcelColumnNameAttribute"/>s from the first of their
/// columns that the sheet has. Each name keeps its own comparison.
/// </summary>
internal sealed class FirstColumnNameReaderFactory(params IReadOnlyList<ColumnNameReaderFactory> factories) : ICellReaderFactory, IColumnNameProviderCellReaderFactory
{
    public IReadOnlyList<ColumnNameReaderFactory> Factories { get; } = factories;

    public ICellReader? GetCellReader(ExcelSheet sheet)
    {
        for (int i = 0; i < Factories.Count; i++)
        {
            if (Factories[i].GetCellReader(sheet) is { } reader)
            {
                return reader;
            }
        }

        return null;
    }

    // For errors: the first name.
    public string GetColumnName(ExcelSheet sheet) => Factories[0].ColumnName;
}

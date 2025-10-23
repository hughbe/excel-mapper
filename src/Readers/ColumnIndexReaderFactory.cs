namespace ExcelMapper.Readers;

/// <summary>
/// Reads the value of a single cell given the zero-based index of it's column.
/// </summary>
public sealed class ColumnIndexReaderFactory : ICellReaderFactory, IColumnIndexProviderCellReaderFactory
{
    /// <summary>
    /// The zero-based index of the column to read.
    /// </summary>
    public int ColumnIndex { get; }

    /// <summary>
    /// Constructs a reader that reads the value of a single cell given the zero-based index of it's column.
    /// </summary>
    /// <param name="columnIndex">The zero-based index of the column to read.</param>
    public ColumnIndexReaderFactory(int columnIndex)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(columnIndex);
        ColumnIndex = columnIndex;
    }

    /// <inheritdoc/>
    public ICellReader? GetCellReader(ExcelSheet sheet)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        if (ColumnIndex >= sheet.NumberOfColumns)
        {
            return null;
        }

        return new ColumnIndexReader(ColumnIndex);
    }

    /// <inheritdoc/>
    public int GetColumnIndex(ExcelSheet sheet) => ColumnIndex;
}

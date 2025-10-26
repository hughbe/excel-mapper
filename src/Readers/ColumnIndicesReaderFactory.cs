namespace ExcelMapper.Readers;

/// <summary>
/// Reads a multiple values of one or more columns given the name of each column.
/// </summary>
public sealed class ColumnIndicesReaderFactory : ICellReaderFactory, ICellsReaderFactory, IColumnIndicesProviderCellReaderFactory
{
    /// <summary>
    /// Gets the zero-based indices for each column to read.
    /// </summary>
    public IReadOnlyList<int> ColumnIndices { get; }

    /// <summary>
    /// Constructs a reader that reads the values of one or more columns with a given zero-based
    /// index and returns the string value of for each column.
    /// </summary>
    /// <param name="columnIndices">The list of zero-based column indices to read.</param>
    public ColumnIndicesReaderFactory(params IReadOnlyList<int> columnIndices)
    {
        ColumnUtilities.ValidateColumnIndices(columnIndices);
        ColumnIndices = columnIndices;
    }

    /// <inheritdoc/>
    public ICellReader? GetCellReader(ExcelSheet sheet)
    {
        ArgumentNullException.ThrowIfNull(sheet);

        foreach (var columnIndex in ColumnIndices)
        {
            if (columnIndex < sheet.NumberOfColumns)
            {
                return new ColumnIndexReader(columnIndex);
            }
        }

        return null;
    }

    /// <inheritdoc/>
    public ICellsReader? GetCellsReader(ExcelSheet sheet)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        foreach (var columnIndex in ColumnIndices)
        {
            if (columnIndex >= sheet.NumberOfColumns)
            {
                return null;
            }
        }

        return new ColumnIndicesReader(ColumnIndices);
    }

    /// <inheritdoc/>
    public IReadOnlyList<int> GetColumnIndices(ExcelSheet sheet) => ColumnIndices;
}

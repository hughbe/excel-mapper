namespace ExcelMapper.Readers;

/// <summary>
/// Reads a multiple values of one or more columns given the name of each column.
/// </summary>
internal sealed class ColumnIndicesReaderFactory : ICellReaderFactory, ICellsReaderFactory, IColumnIndicesProviderCellReaderFactory
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
        ThrowHelpers.ThrowIfNull(sheet, nameof(sheet));

        foreach (var columnIndex in ColumnIndices)
        {
            var sheetColumnIndex = sheet.ToSheetColumn(columnIndex);
            if (sheetColumnIndex != -1)
            {
                return new ColumnIndexReader(sheetColumnIndex);
            }
        }

        return null;
    }

    /// <inheritdoc/>
    public ICellsReader? GetCellsReader(ExcelSheet sheet)
    {
        ThrowHelpers.ThrowIfNull(sheet, nameof(sheet));
        var sheetColumnIndices = new int[ColumnIndices.Count];
        for (int i = 0; i < sheetColumnIndices.Length; i++)
        {
            sheetColumnIndices[i] = sheet.ToSheetColumn(ColumnIndices[i]);
            if (sheetColumnIndices[i] == -1)
            {
                return null;
            }
        }

        return new ColumnIndicesReader(sheetColumnIndices);
    }

    /// <inheritdoc/>
    public IReadOnlyList<int> GetColumnIndices(ExcelSheet sheet) => ColumnIndices;
}

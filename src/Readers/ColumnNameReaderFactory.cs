namespace ExcelMapper.Readers;

/// <summary>
/// Reads the value of a single cell given the name of it's column.
/// </summary>
public sealed class ColumnNameReaderFactory : ICellReaderFactory, IColumnNameProviderCellReaderFactory
{
    /// <summary>
    /// The name of the column to read.
    /// </summary>
    public string ColumnName { get; }

    /// <summary>
    /// The string comparison to use when matching column names.
    /// </summary>
    public StringComparison Comparison { get; }

    /// <summary>
    /// Constructs a reader that reads the value of a single cell given the name of it's column.
    /// </summary>
    /// <param name="columnName">The name of the column to read.</param>
    /// <param name="comparison">The string comparison to use when matching column names.</param>
    public ColumnNameReaderFactory(string columnName, StringComparison comparison = StringComparison.OrdinalIgnoreCase)
    {
        ColumnUtilities.ValidateColumnName(columnName);
        EnumUtilities.ValidateIsDefined(comparison);
        ColumnName = columnName;
        Comparison = comparison;
    }

    /// <inheritdoc/>
    public ICellReader? GetCellReader(ExcelSheet sheet)
    {
        if (ColumnUtilities.TryGetColumnIndex(sheet, ColumnName, Comparison, out var columnIndex))
        {
            return new ColumnIndexReader(columnIndex);
        }

        return null;
    }

    /// <inheritdoc/>
    public string GetColumnName(ExcelSheet sheet) => ColumnName;
}

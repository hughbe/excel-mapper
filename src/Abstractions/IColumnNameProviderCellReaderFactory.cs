namespace ExcelMapper;

/// <summary>
/// Provides the column name for a given sheet.
/// </summary>
public interface IColumnNameProviderCellReaderFactory
{
    /// <summary>
    /// Gets the column name for the given sheet.
    /// </summary>
    /// <param name="sheet">The sheet to get the column name for.</param>
    /// <returns>The column name for the given sheet.</returns>
    string? GetColumnName(ExcelSheet sheet);
}

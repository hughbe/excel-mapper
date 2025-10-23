namespace ExcelMapper;

/// <summary>
/// Provides the column names for a given sheet.
/// </summary>
public interface IColumnNamesProviderCellReaderFactory
{
    /// <summary>
    /// Gets the column names for the given sheet.
    /// </summary>
    /// <param name="sheet">The sheet to get the column names for.</param>
    /// <returns>The column names for the given sheet, or null if no column names are available.</returns>
    string[]? GetColumnNames(ExcelSheet sheet);
}

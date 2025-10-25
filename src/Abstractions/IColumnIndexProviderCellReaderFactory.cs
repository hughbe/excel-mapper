namespace ExcelMapper;

/// <summary>
/// Provides the column index for a given sheet.
/// This is a performance optimization to cache column indices.
/// </summary>
public interface IColumnIndexProviderCellReaderFactory
{
    /// <summary>
    /// Gets the column index for the given sheet.
    /// </summary>
    /// <param name="sheet">The sheet to get the column index for.</param>
    /// <returns>The column index for the given sheet.</returns>
    int? GetColumnIndex(ExcelSheet sheet);
}

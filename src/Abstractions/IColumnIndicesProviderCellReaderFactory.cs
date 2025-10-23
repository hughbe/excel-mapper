namespace ExcelMapper;

/// <summary>
/// Provides the column indices for a given sheet.
/// This is a performance optimization to cache column indices.
/// </summary>
public interface IColumnIndicesProviderCellReaderFactory
{
    /// <summary>
    /// Gets the column indices for the given sheet.
    /// </summary>
    /// <param name="sheet">The sheet to get the column indices for.</param>
    /// <returns>The column indices for the given sheet, or null if no column indices are available.</returns>
    int[]? GetColumnIndices(ExcelSheet sheet);
}

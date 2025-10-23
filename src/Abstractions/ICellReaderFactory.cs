namespace ExcelMapper.Abstractions;

/// <summary>
/// Factory for getting cell readers for a given sheet.
/// This is a performance optimization to cache column indices.
/// </summary>
public interface ICellReaderFactory
{
    /// <summary>
    /// Gets a cell reader for the given sheet.
    /// </summary>
    /// <param name="sheet">The sheet to get the cell reader for.</param>
    /// <returns>The cell reader for the given sheet, or null if no cell reader is available.</returns>
    ICellReader? GetCellReader(ExcelSheet sheet);
}

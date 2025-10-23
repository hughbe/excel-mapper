namespace ExcelMapper.Abstractions;

/// <summary>
/// Factory for getting cells readers for a given sheet.
/// This is a performance optimization to cache column indices.
/// </summary>
public interface ICellsReaderFactory
{
    /// <summary>
    /// Gets a cells reader for the given sheet.
    /// </summary>
    ICellsReader? GetCellsReader(ExcelSheet sheet);
}

namespace ExcelMapper.Abstractions;

/// <summary>
/// An interface that creates a reader for multiple cells on a specific sheet.
/// This is a performance optimisation to avoid recalculating
/// column indices for each row.
/// </summary>
public interface ICellsReaderFactory
{
    ICellsReader? GetCellsReader(ExcelSheet sheet);
}

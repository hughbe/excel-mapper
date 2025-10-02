namespace ExcelMapper.Abstractions;

/// <summary>
/// An interface that creates a reader for a cell on a specific sheet.
/// This is a performance optimisation to avoid recalculating
/// column indices for each row.
/// </summary>
public interface ICellReaderFactory
{
    ICellReader? GetReader(ExcelSheet sheet);
}

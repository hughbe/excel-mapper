namespace ExcelMapper.Abstractions;

public interface IReadCellResultEnumerator : IEnumerator<ReadCellResult>
{
    /// <summary>
    /// Gets the number of cells to read for the given sheet.
    /// </summary>
    public int Count { get; }
}
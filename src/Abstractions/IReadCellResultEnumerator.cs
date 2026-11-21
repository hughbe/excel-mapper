namespace ExcelMapper.Abstractions;

internal interface IReadCellResultEnumerator : IEnumerator<ReadCellResult>
{
    /// <summary>
    /// Gets the number of cells to read for the given sheet.
    /// </summary>
    public int Count { get; }
}
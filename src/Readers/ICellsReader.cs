using ExcelDataReader;

namespace ExcelMapper.Readers;

/// <summary>
/// Reads the cells of a collection or dictionary from the current row: several columns
/// (<see cref="ColumnIndicesReader"/>), or the values of one cell split by separators
/// (<see cref="SplitCellsReader"/>). A reader is made once per sheet and member (see
/// <see cref="SheetView.GetCellsReader(Columns)"/>) and reused for every row, so it isn't shared
/// between threads.
/// </summary>
internal interface ICellsReader
{
    /// <summary>
    /// Starts reading the current row's cells.
    /// </summary>
    /// <param name="reader">The reader, on the row being read.</param>
    /// <param name="preserveFormatting">Whether cells are read as the text Excel shows.</param>
    /// <param name="count">The number of cells, so the collection can be created at its size.</param>
    /// <returns>False if the row doesn't have the cell to split.</returns>
    bool Start(IExcelDataReader reader, bool preserveFormatting, out int count);

    /// <summary>
    /// Reads the next cell.
    /// </summary>
    /// <returns>False if there are no more cells.</returns>
    bool TryGetNext(out ReadCellResult result);

    /// <summary>
    /// Finishes reading the row, so the reader holds no values from it.
    /// </summary>
    void Reset();
}

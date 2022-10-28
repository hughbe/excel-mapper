namespace ExcelMapper.Abstractions;
 
/// <summary>
/// Metadata about the output of reading the value of a single cell.
/// </summary>
public struct ExcelCell
{
    /// <summary>
    /// The sheet that contains the cell.
    /// </summary>
    public ExcelSheet Sheet { get; }

    /// <summary>
    /// The index of the row that contains the cell.
    /// </summary>
    public int RowIndex { get; }

    /// <summary>
    /// The index of the column that contains the cell.
    /// </summary>
    public int ColumnIndex { get; }

    /// <summary>
    /// Constructs an object describing the output of reading the value of a single cell.
    /// </summary>
    /// <param name="sheet">The sheet that contains the cell.</param>
    /// <param name="rowIndex">The index of the row that contains the cell.</param>
    /// <param name="columnIndex">The index of the column that contains the cell.</param>
    /// <param name="value">The value of thecell.</param>
    public ExcelCell(ExcelSheet sheet, int rowIndex, int columnIndex)
    {
        Sheet = sheet;
        RowIndex = rowIndex;
        ColumnIndex = columnIndex;
    }
}

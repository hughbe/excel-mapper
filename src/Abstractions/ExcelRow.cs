namespace ExcelMapper.Abstractions;
 
/// <summary>
/// Metadata about the context of reading a row within a sheet.
/// </summary>
public struct ExcelRow
{
    /// <summary>
    /// The sheet that contains the row.
    /// </summary>
    public ExcelSheet Sheet { get; }

    /// <summary>
    /// The index of the row.
    /// </summary>
    public int RowIndex { get; }

    /// <summary>
    /// The number of columns in the row.
    /// </summary>
    public int ColumnCount { get; }

    /// <summary>
    /// Metadata about the context of reading a row within a sheet.
    /// </summary>
    /// <param name="sheet">The sheet that contains the row.</param>
    /// <param name="rowIndex">The index of the row.</param>
    public ExcelRow(ExcelSheet sheet, int rowIndex, int columnCount)
    {
        Sheet = sheet;
        RowIndex = rowIndex;
        ColumnCount = columnCount;
    }
}

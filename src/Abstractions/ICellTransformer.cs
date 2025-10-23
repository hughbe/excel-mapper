namespace ExcelMapper.Abstractions;

/// <summary>
/// Transforms the string value of a cell read from Excel.
/// </summary>
public interface ICellTransformer
{
    /// <summary>
    /// Transforms the string value of a cell read from Excel.
    /// </summary>
    /// <param name="sheet">The sheet the cell is on.</param>
    /// <param name="rowIndex">The row index of the cell.</param>
    /// <param name="readResult">The read result of the cell.</param>
    /// <returns>The transformed string value.</returns>
    string? TransformStringValue(ExcelSheet sheet, int rowIndex, ReadCellResult readResult);
}

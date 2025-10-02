namespace ExcelMapper.Abstractions;

/// <summary>
/// Metadata about the output of reading the value of a single cell.
/// </summary>
public readonly struct ReadCellResult
{
    /// <summary>
    /// The index of the column that contains the cell.
    /// </summary>
    public int ColumnIndex { get; }

    /// <summary>
    /// The string value of the cell.
    /// </summary>
    public string? StringValue { get; }

    /// <summary>
    /// Constructs an object describing the output of reading the value of a single cell.
    /// </summary>
    /// <param name="columnIndex">The index of the column that contains the cell.</param>
    /// <param name="stringValue">The string value of the cell.</param>
    public ReadCellResult(int columnIndex, string? stringValue)
    {
        ColumnIndex = columnIndex;
        StringValue = stringValue;
    }
}

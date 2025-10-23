using ExcelDataReader;

namespace ExcelMapper.Readers;

/// <summary>
/// Reads the value of a cell.
/// </summary>
public class ColumnIndexReader : ICellReader
{
    /// <summary>
    /// The index of the column to read.
    /// </summary>
    public int ColumnIndex { get; }

    /// <summary>
    /// Constructs a reader that reads the value of a cell at the given column index.
    /// </summary>
    /// <param name="columnIndex">The index of the column to read.</param>
    public ColumnIndexReader(int columnIndex)
    {
        ColumnUtilities.ValidateColumnIndex(columnIndex, nameof(columnIndex));
        ColumnIndex = columnIndex;
    }

    /// <inheritdoc/>
    public bool TryGetValue(IExcelDataReader reader, bool preserveFormatting, out ReadCellResult result)
    {
        if (ColumnIndex >= reader.FieldCount)
        {
            result = default;
            return false;
        }

        result = new ReadCellResult(ColumnIndex, reader, preserveFormatting);
        return true;
    }
}

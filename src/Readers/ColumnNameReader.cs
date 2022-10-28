using ExcelDataReader;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Readers;

/// <summary>
/// Reads the value of a single cell given the name of it's column.
/// </summary>
public sealed class ColumnNameValueReader : ICellReader
{
    /// <summary>
    /// The name of the column to read.
    /// </summary>
    public string ColumnName { get; }

    /// <summary>
    /// Constructs a reader that reads the value of a single cell given the name of it's column.
    /// </summary>
    /// <param name="columnName">The name of the column to read.</param>
    public ColumnNameValueReader(string columnName)
    {
        if (columnName == null)
        {
            throw new ArgumentNullException(nameof(columnName));
        }

        if (columnName.Length == 0)
        {
            throw new ArgumentException("Column name cannot be empty.", nameof(columnName));
        }

        ColumnName = columnName;
    }

    public bool TryGetCell(ExcelRow row, out ExcelCell cell)
    {
        if (row.Sheet.Heading == null)
        {
            throw new ExcelMappingException($"The sheet \"{row.Sheet.Name}\" does not have a heading. Use a column index mapping instead.");
        }

        if (!row.Sheet.Heading.TryGetColumnIndex(ColumnName, out int index))
        {
            cell = default;
            return false;
        }

        cell = new ExcelCell(row.Sheet, row.RowIndex, index);
        return true;
    }
}

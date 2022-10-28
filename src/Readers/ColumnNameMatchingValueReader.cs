using ExcelDataReader;
using ExcelMapper;
using ExcelMapper.Abstractions;

/// <summary>
/// Reads the value of a single cell given the predicate matching the column name.
/// </summary>
public sealed class ColumnNameMatchingValueReader : ICellReader
{
    private readonly Func<string, bool> _predicate;

    /// <summary>
    /// Constructs a reader that reads the value of a single cell given the predicate matching the column name.
    /// </summary>
    /// <param name="predicate">The predicate containing the column name to read.</param>
    public ColumnNameMatchingValueReader(Func<string, bool> predicate)
    {
        _predicate = predicate ?? throw new ArgumentNullException(nameof(predicate));
    }

    public bool TryGetCell(ExcelRow row, out ExcelCell cell)
    {
        if (row.Sheet.Heading == null)
        {
            throw new ExcelMappingException($"The sheet \"{row.Sheet.Name}\" does not have a heading. Use a column index mapping instead.");
        }

        if (!row.Sheet.Heading.TryGetFirstColumnMatchingIndex(_predicate, out int index))
        {
            cell = default;
            return false;
        }

        cell = new ExcelCell(row.Sheet, row.RowIndex, index);
        return true;
    }
}
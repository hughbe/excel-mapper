using ExcelDataReader;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Readers
{
    /// <summary>
    /// Finds a cell from a zero-based column index.
    /// </summary>
    public sealed class ColumnIndexValueReader : ICellReader
    {
        /// <summary>
        /// The zero-based index of the column to read.
        /// </summary>
        public int ColumnIndex { get; }

        /// <summary>
        /// Constructs a reader that reads the value of a single cell given a zero-based column index.
        /// </summary>
        /// <param name="columnIndex">The zero-based index of the column to read.</param>
        public ColumnIndexValueReader(int columnIndex)
        {
            if (columnIndex < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(columnIndex), columnIndex, $"Column index {columnIndex} must be greater or equal to zero.");
            }

            ColumnIndex = columnIndex;
        }

        public bool TryGetCell(ExcelRow row, out ExcelCell cell)
        {
            if (ColumnIndex >= row.ColumnCount)
            {
                cell = default;
                return false;
            }

            cell = new ExcelCell(row.Sheet, row.RowIndex, ColumnIndex);
            return true;
        }
    }
}

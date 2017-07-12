using System;
using ExcelDataReader;

namespace ExcelMapper.Mappings.Readers
{
    /// <summary>
    /// Reads a single value of a column given the zero-based index of that column.
    /// </summary>
    public sealed class ColumnIndexReader : ISingleValueReader
    {
        public int ColumnIndex { get; }

        public ColumnIndexReader(int columnIndex)
        {
            if (columnIndex < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(columnIndex), columnIndex, $"Column index {columnIndex} must be greater or equal to zero.");
            }

            ColumnIndex = columnIndex;
        }

        public ReadResult GetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader)
        {
            return new ReadResult(ColumnIndex, reader.GetString(ColumnIndex));
        }
    }
}

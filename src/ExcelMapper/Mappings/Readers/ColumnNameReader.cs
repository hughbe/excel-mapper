using System;
using ExcelDataReader;

namespace ExcelMapper.Mappings.Readers
{
    /// <summary>
    /// Reads a single value of a column given the name of the column.
    /// </summary>
    public sealed class ColumnNameReader : ISingleValueReader
    {
        public string ColumnName { get; }

        public ColumnNameReader(string columnName)
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

        public ReadResult GetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader)
        {
            int index = sheet.Heading.GetColumnIndex(ColumnName);
            return new ReadResult(index, reader.GetString(index));
        }
    }
}

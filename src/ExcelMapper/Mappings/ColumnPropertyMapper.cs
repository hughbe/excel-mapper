using System;
using ExcelDataReader;

namespace ExcelMapper.Mappings
{
    internal sealed class ColumnPropertyMapper : ISinglePropertyMapper
    {
        public string ColumnName { get; }

        internal ColumnPropertyMapper(string columnName)
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

        public MapResult GetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader)
        {
            int index = sheet.Heading.GetColumnIndex(ColumnName);
            return new MapResult(index, reader.GetString(index));
        }
    }
}

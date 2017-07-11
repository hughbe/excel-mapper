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
                throw new ArgumentException(nameof(columnName));
            }

            ColumnName = columnName;
        }

        public int GetColumnIndex(ExcelSheet sheet, int rowIndex, IExcelDataReader reader)
        {
            return sheet.Heading.GetColumnIndex(ColumnName);
        }
    }
}

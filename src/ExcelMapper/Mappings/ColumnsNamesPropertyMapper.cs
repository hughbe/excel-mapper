using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDataReader;
using ExcelMapper.Utilities;

namespace ExcelMapper.Mappings
{
    public sealed class ColumnsNamesPropertyMapper : IMultiPropertyMapper
    {
        public string[] ColumnNames { get; }

        public ColumnsNamesPropertyMapper(string[] columnNames)
        {
            if (columnNames == null)
            {
                throw new ArgumentNullException(nameof(columnNames));
            }

            if (columnNames.Length == 0)
            {
                throw new ArgumentException("Column names cannot be empty.", nameof(columnNames));
            }
            
            foreach (string columnName in columnNames)
            {
                if (columnName == null)
                {
                    throw new ArgumentException($"Null column name in {columnNames.ArrayJoin()}.", nameof(columnNames));
                }
            }

            ColumnNames = columnNames;
        }

        public IEnumerable<MapResult> GetValues(ExcelSheet sheet, int rowIndex, IExcelDataReader reader)
        {
            return ColumnNames.Select(columnName =>
            {
                int index = sheet.Heading.GetColumnIndex(columnName);
                return new MapResult(index, reader.GetString(index));
            });
        }
    }
}

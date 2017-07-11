using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDataReader;
using ExcelMapper.Utilities;

namespace ExcelMapper.Mappings
{
    internal sealed class ColumnsPropertyMapper : IMultiPropertyMapper
    {
        public string[] ColumnNames { get; }

        internal ColumnsPropertyMapper(IEnumerable<string> columnNames)
        {
            if (columnNames == null)
            {
                throw new ArgumentNullException(nameof(columnNames));
            }

            foreach (string columnName in columnNames)
            {
                if (columnNames == null)
                {
                    throw new ArgumentException($"Null column name in {columnNames.ArrayJoin()}.", nameof(columnNames));
                }
            }

            ColumnNames = columnNames.ToArray();
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

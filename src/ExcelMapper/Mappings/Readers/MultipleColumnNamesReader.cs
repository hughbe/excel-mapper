using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDataReader;
using ExcelMapper.Utilities;

namespace ExcelMapper.Mappings.Readers
{
    /// <summary>
    /// Reads a multiple values of one or more columns given the zero-based index of each column.
    /// </summary>
    public sealed class MultipleColumnNamesReader : IMultipleValuesReader
    {
        public string[] ColumnNames { get; }

        public MultipleColumnNamesReader(string[] columnNames)
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

        public IEnumerable<ReadResult> GetValues(ExcelSheet sheet, int rowIndex, IExcelDataReader reader)
        {
            return ColumnNames.Select(columnName =>
            {
                int index = sheet.Heading.GetColumnIndex(columnName);
                return new ReadResult(index, reader.GetString(index));
            });
        }
    }
}

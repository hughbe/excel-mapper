using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDataReader;
using ExcelMapper.Abstractions;
using ExcelMapper.Utilities;

namespace ExcelMapper.Readers
{
    /// <summary>
    /// Reads a multiple values of one or more columns given the name of each column.
    /// </summary>
    public sealed class MultipleColumnNamesValueReader : IMultipleCellValuesReader
    {
        /// <summary>
        /// Gets the names of each column to read.
        /// </summary>
        public string[] ColumnNames { get; }

        /// <summary>
        /// Constructs a reader that reads the values of one or more columns with a given name
        /// and returns the string value of for each column.
        /// </summary>
        /// <param name="columnNames">The names of each column to read.</param>
        public MultipleColumnNamesValueReader(params string[] columnNames)
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

        public bool TryGetValues(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, out IEnumerable<ReadCellValueResult> result)
        {
            if (sheet.Heading == null)
            {
                throw new ExcelMappingException($"The sheet \"{sheet.Name}\" does not have a heading. Use a column index mapping instead.");
            }

            var values = new ReadCellValueResult[ColumnNames.Length];
            for (int i = 0; i < ColumnNames.Length; i++)
            {
                if (!sheet.Heading.TryGetColumnIndex(ColumnNames[i], out int index))
                {
                    result = default;
                    return false;
                }

                var value = reader[index]?.ToString();
                values[i] = new ReadCellValueResult(index, value);
            }

            result = values;
            return true;
        }
    }
}

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
    public sealed class MultipleColumnIndicesValueReader : IMultipleCellValuesReader
    {
        /// <summary>
        /// Gets the zero-based index of each column to read.
        /// </summary>
        public int[] ColumnIndices { get; }

        /// <summary>
        /// Constructs a reader that reads the values of one or more columns with a given zero-based
        /// index and returns the string value of for each column.
        /// </summary>
        /// <param name="columnIndices">The list of zero-based column indices to read.</param>
        public MultipleColumnIndicesValueReader(int[] columnIndices)
        {
            if (columnIndices == null)
            {
                throw new ArgumentNullException(nameof(columnIndices));
            }

            if (columnIndices.Length == 0)
            {
                throw new ArgumentException("Column indices cannot be empty.", nameof(columnIndices));
            }

            foreach (int columnIndex in columnIndices)
            {
                if (columnIndex < 0)
                {
                    throw new ArgumentOutOfRangeException(nameof(columnIndices), columnIndex, $"Negative column index in {columnIndices.ArrayJoin()}.");
                }
            }

            ColumnIndices = columnIndices;
        }

        public bool TryGetValues(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, out IEnumerable<ReadCellValueResult> result)
        {
            result = ColumnIndices.Select(columnIndex =>
            {
                var value = reader[columnIndex]?.ToString();
                return new ReadCellValueResult(columnIndex, value);
            });
            return true;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDataReader;
using ExcelMapper.Utilities;

namespace ExcelMapper.Mappings.Readers
{
    /// <summary>
    /// Reads a multiple values of one or more columns given the name of each column.
    /// </summary>
    public sealed class MultipleColumnIndicesReader : IMultipleValuesReader
    {
        public int[] ColumnIndices { get; }

        public MultipleColumnIndicesReader(int[] columnIndices)
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

        public IEnumerable<ReadResult> GetValues(ExcelSheet sheet, int rowIndex, IExcelDataReader reader)
        {
            return ColumnIndices.Select(i => new ReadResult(i, reader.GetString(i)));
        }
    }
}

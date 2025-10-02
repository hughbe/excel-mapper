using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using ExcelDataReader;
using ExcelMapper.Abstractions;
using ExcelMapper.Utilities;

namespace ExcelMapper.Readers;

/// <summary>
/// Reads the value of a cell.
/// </summary>
public class ColumnIndicesReader : ICellsReader
{
    /// <summary>
    /// Gets the zero-based indices for each column to read.
    /// </summary>
    public int[] ColumnIndices { get; }

    public ColumnIndicesReader(int[] columnIndices)
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

    public bool TryGetValues(IExcelDataReader reader, [NotNullWhen(true)] out IEnumerable<ReadCellResult>? result)
    {
        result = ColumnIndices.Select(columnIndex =>
        {
            var value = reader[columnIndex]?.ToString();
            return new ReadCellResult(columnIndex, value);
        });
        return true;
    }
}

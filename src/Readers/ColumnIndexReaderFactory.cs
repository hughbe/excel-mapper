using System;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Readers;

/// <summary>
/// Reads the value of a single cell given the zero-based index of it's column.
/// </summary>
public sealed class ColumnIndexReaderFactory : ICellReaderFactory
{
    /// <summary>
    /// The zero-based index of the column to read.
    /// </summary>
    public int ColumnIndex { get; }

    /// <summary>
    /// Constructs a reader that reads the value of a single cell given the zero-based index of it's column.
    /// </summary>
    /// <param name="columnIndex">The zero-based index of the column to read.</param>
    public ColumnIndexReaderFactory(int columnIndex)
    {
        if (columnIndex < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(columnIndex), columnIndex, $"Column index {columnIndex} must be greater or equal to zero.");
        }

        ColumnIndex = columnIndex;
    }

    public ICellReader? GetCellReader(ExcelSheet sheet)
    {
        if (sheet == null)
        {
            throw new ArgumentNullException(nameof(sheet));
        }
        if (ColumnIndex >= sheet.NumberOfColumns)
        {
            return null;
        }

        return new ColumnIndexReader(ColumnIndex);
    }
}

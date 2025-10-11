using System;
using ExcelMapper.Abstractions;
using ExcelMapper.Utilities;

namespace ExcelMapper.Readers;

/// <summary>
/// Reads a multiple values of one or more columns given the name of each column.
/// </summary>
public sealed class ColumnIndicesReaderFactory : ICellReaderFactory, ICellsReaderFactory, IColumnIndicesProviderCellReaderFactory
{
    /// <summary>
    /// Gets the zero-based indices for each column to read.
    /// </summary>
    public int[] ColumnIndices { get; }

    /// <summary>
    /// Constructs a reader that reads the values of one or more columns with a given zero-based
    /// index and returns the string value of for each column.
    /// </summary>
    /// <param name="columnIndices">The list of zero-based column indices to read.</param>
    public ColumnIndicesReaderFactory(params int[] columnIndices)
    {
        ColumnUtilities.ValidateColumnIndices(columnIndices, nameof(columnIndices));
        ColumnIndices = columnIndices;
    }

    public ICellReader? GetCellReader(ExcelSheet sheet)
    {
        if (sheet == null)
        {
            throw new ArgumentNullException(nameof(sheet));
        }

        foreach (int columnIndex in ColumnIndices)
        {
            if (columnIndex < sheet.NumberOfColumns)
            {
                return new ColumnIndexReader(columnIndex);
            }
        }

        return null;
    }

    public ICellsReader? GetCellsReader(ExcelSheet sheet)
    {
        if (sheet == null)
        {
            throw new ArgumentNullException(nameof(sheet));
        }
        foreach (int columnIndex in ColumnIndices)
        {
            if (columnIndex >= sheet.NumberOfColumns)
            {
                return null;
            }
        }

        return new ColumnIndicesReader(ColumnIndices);
    }

    public int[] GetColumnIndices(ExcelSheet sheet) => ColumnIndices;
}

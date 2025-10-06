using System;
using ExcelDataReader;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Readers;

/// <summary>
/// Reads the value of a cell.
/// </summary>
public class ColumnIndexReader : ICellReader
{
    public int ColumnIndex { get; }

    public ColumnIndexReader(int columnIndex)
    {
        if (columnIndex < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(columnIndex), columnIndex, $"Column index {columnIndex} must be greater or equal to zero.");
        }

        ColumnIndex = columnIndex;
    }

    public bool TryGetValue(IExcelDataReader reader, bool preserveFormatting, out ReadCellResult result)
    {
        if (ColumnIndex >= reader.FieldCount)
        {
            result = default;
            return false;
        }

        result = new ReadCellResult(ColumnIndex, reader, preserveFormatting);
        return true;
    }
}

using System;
using ExcelDataReader;
using ExcelMapper.Abstractions;
using ExcelMapper.Utilities;

namespace ExcelMapper.Readers;

/// <summary>
/// Reads the value of a cell.
/// </summary>
public class ColumnIndexReader : ICellReader
{
    public int ColumnIndex { get; }

    public ColumnIndexReader(int columnIndex)
    {
        ColumnUtilities.ValidateColumnIndex(columnIndex, nameof(columnIndex));
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

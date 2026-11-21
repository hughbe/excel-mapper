using ExcelDataReader;

namespace ExcelMapper.Utilities;

/// <summary>
/// Owns the position of an importer's reader. The reader is forward-only and shared by every
/// <see cref="ExcelSheet"/> of the importer, so sheets don't move it themselves: they ask for the
/// sheet and row they need, and this reads forward or rewinds as needed. This is the only code
/// that calls <see cref="IExcelDataReader.Read"/>, <see cref="System.Data.IDataReader.NextResult"/>
/// or <see cref="IExcelDataReader.Reset"/>.
/// </summary>
internal sealed class RowSource
{
    private readonly IExcelDataReader _reader;
    private readonly bool _rowCountExcludesEmptyLines;

    // The sheet and row the reader is on. -1 for the sheet means it isn't known (for example after
    // NextResult failed), so the next move rewinds. -1 for the row means before the first row.
    private int _sheetIndex;
    private int _rowIndex = -1;

    // Whether the reader has passed the last row of the sheet. _rowIndex is then the index of the
    // first row that doesn't exist.
    private bool _ended;

    // The number of rows read from the current sheet that have a value in any column.
    private int _rowsWithValues;

    // Row counts by sheet, for readers whose own count leaves out empty lines.
    private Dictionary<int, int>? _rowCounts;

    public RowSource(IExcelDataReader reader)
    {
        _reader = reader;
        _rowCountExcludesEmptyLines = reader.GetType().FullName == "ExcelDataReader.ExcelCsvReader";
    }

    /// <summary>
    /// Moves the reader to the start of a sheet, before its first row.
    /// </summary>
    /// <returns>False if the workbook doesn't have the sheet.</returns>
    public bool TryMoveToSheet(int sheetIndex)
    {
        if (_sheetIndex == sheetIndex && _rowIndex == -1 && !_ended)
        {
            return true;
        }

        if (_sheetIndex == -1 || _sheetIndex > sheetIndex || (_sheetIndex == sheetIndex && (_rowIndex != -1 || _ended)))
        {
            Rewind();
        }

        while (_sheetIndex < sheetIndex)
        {
            if (!_reader.NextResult())
            {
                // The reader's position is no longer known.
                _sheetIndex = -1;
                return false;
            }

            _sheetIndex++;
            _rowIndex = -1;
            _ended = false;
            _rowsWithValues = 0;
        }

        return true;
    }

    /// <summary>
    /// Moves the reader to a row of a sheet, reading forward or rewinding as needed. A row index of
    /// -1 moves to the start of the sheet.
    /// </summary>
    /// <returns>False if the sheet doesn't have the row. Empty lines at the end of a CSV file aren't
    /// rows.</returns>
    public bool TryMoveTo(int sheetIndex, int rowIndex)
    {
        if (_sheetIndex != sheetIndex || _rowIndex > rowIndex)
        {
            if (!TryMoveToSheet(sheetIndex))
            {
                return false;
            }
        }

        while (_rowIndex < rowIndex)
        {
            if (!ReadNextRow())
            {
                return false;
            }
        }

        return !(_ended && _rowIndex == rowIndex);
    }

    /// <summary>
    /// Gets the index of the last row of the current sheet that the reader has read, or -1 if none.
    /// After a move past the end of the sheet, this is the sheet's last row.
    /// </summary>
    public int LastRowIndex => _ended ? _rowIndex - 1 : _rowIndex;

    /// <summary>
    /// Gets the number of rows in a sheet, not counting empty lines at the end of a CSV file.
    /// </summary>
    public int GetRowCount(int sheetIndex)
    {
        // The CSV reader's row count leaves out empty lines, so count the rows by reading them,
        // once per sheet. Callers move to the row they need before reading it, so the position
        // doesn't need restoring.
        if (!_rowCountExcludesEmptyLines)
        {
            // The row count is for the sheet the reader is on.
            if (_sheetIndex != sheetIndex)
            {
                TryMoveToSheet(sheetIndex);
            }

            return _reader.RowCount;
        }

        _rowCounts ??= [];
        if (!_rowCounts.TryGetValue(sheetIndex, out var rowCount))
        {
            TryMoveTo(sheetIndex, -1);
            while (ReadNextRow())
            {
            }

            rowCount = _rowIndex;
            _rowCounts[sheetIndex] = rowCount;
        }

        return rowCount;
    }

    /// <summary>
    /// Whether the cells of the current row in the given columns are all empty.
    /// </summary>
    public bool IsBlankRow(int startColumnIndex, int endColumnIndex)
    {
        for (int i = startColumnIndex; i < endColumnIndex; i++)
        {
            var value = _reader.GetValue(i);
            if (value is not null && !(value is string stringValue && string.IsNullOrEmpty(stringValue)))
            {
                return false;
            }
        }

        return true;
    }

    private void Rewind()
    {
        _reader.Reset();
        _sheetIndex = 0;
        _rowIndex = -1;
        _ended = false;
        _rowsWithValues = 0;
    }

    private bool ReadNextRow()
    {
        if (_ended)
        {
            return false;
        }

        if (!_reader.Read())
        {
            _rowIndex++;
            _ended = true;
            return false;
        }

        _rowIndex++;
        if (!IsBlankRow(0, _reader.FieldCount))
        {
            _rowsWithValues++;
            return true;
        }

        // The CSV reader's row count leaves out empty lines, so it can be less than the number of
        // rows read. Every row with a value is counted, so once that many rows with values have
        // been read, only empty lines remain: treat them as the end of the sheet. For other readers
        // the row count includes every row, so this is never true.
        if (_rowsWithValues >= _reader.RowCount)
        {
            _ended = true;
            return false;
        }

        return true;
    }
}

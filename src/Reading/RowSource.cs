using ExcelDataReader;

namespace ExcelMapper.Reading;

/// <summary>
/// Owns the position of an importer's reader. The reader is forward-only and shared by every
/// <see cref="ExcelSheet"/> of the importer, so sheets don't move it themselves: they ask for the
/// sheet and row they need, and this reads forward or rewinds as needed. This is the only code
/// that calls <see cref="System.Data.IDataReader.Read"/>, <see cref="System.Data.IDataReader.NextResult"/>
/// or <see cref="IExcelDataReader.Reset"/>.
/// </summary>
internal sealed class RowSource
{
    private readonly IExcelDataReader _reader;

    // Reads blank rows that the reader has read past (see ReadNextRow), created when first needed.
    private BlankRowReader? _blankRowReader;

    // The sheet and row the reader is on. -1 for the sheet means it isn't known (for example after
    // NextResult failed), so the next move rewinds. -1 for the row means before the first row.
    private int _sheetIndex;
    private int _rowIndex = -1;

    // The row the reader itself is on. It can be ahead of _rowIndex: to tell whether a blank row
    // is at the end of the sheet, the reader reads ahead to the next row with a value, and the
    // blank rows from _rowIndex up to it are read with a BlankRowReader (see ReadNextRow).
    private int _readerRowIndex = -1;

    // Whether the reader has passed the last row of the sheet. _rowIndex is then the index of the
    // first row that doesn't exist.
    private bool _ended;

    // Row counts by sheet, not counting blank rows at the end.
    private Dictionary<int, int>? _rowCounts;

    public RowSource(IExcelDataReader reader)
    {
        _reader = reader;
    }

    /// <summary>
    /// Gets the reader of the current row's cells: the reader itself, or, for a blank row it has
    /// read past, a reader whose cells are all empty.
    /// </summary>
    public IExcelDataReader CurrentReader => _rowIndex == _readerRowIndex ? _reader : (_blankRowReader ??= new BlankRowReader(_reader));

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
            _readerRowIndex = -1;
            _ended = false;
        }

        return true;
    }

    /// <summary>
    /// Moves the reader to a row of a sheet, reading forward or rewinding as needed. A row index of
    /// -1 moves to the start of the sheet.
    /// </summary>
    /// <returns>False if the sheet doesn't have the row. Blank rows at the end of a sheet, such as
    /// empty lines at the end of a CSV file or formatted rows below a table, aren't rows.</returns>
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
    /// Gets the number of rows in a sheet, not counting blank rows at the end.
    /// </summary>
    public int GetRowCount(int sheetIndex)
    {
        // The readers' own counts include blank rows at the end (except the CSV reader, which
        // leaves out every empty line), so count the rows by reading them, once per sheet. Callers
        // move to the row they need before reading it, so the position doesn't need restoring.
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
            var value = CurrentReader.GetValue(i);
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
        _readerRowIndex = -1;
        _ended = false;
    }

    private bool ReadNextRow()
    {
        if (_ended)
        {
            return false;
        }

        // A blank row the reader read past while reading ahead.
        if (_rowIndex < _readerRowIndex)
        {
            _rowIndex++;
            return true;
        }

        _rowIndex++;
        if (!ReadReaderRow())
        {
            _ended = true;
            return false;
        }

        if (!IsReaderRowBlank())
        {
            return true;
        }

        // A blank row is only a row if a row with a value follows it, so that blank rows at the end
        // of a sheet (formatted rows below a table, or empty lines at the end of a CSV file) aren't
        // read as rows. Read ahead to the next row with a value: the blank rows before it are read
        // with a BlankRowReader. Blank rows are read once either way, so this costs nothing extra.
        while (ReadReaderRow())
        {
            if (!IsReaderRowBlank())
            {
                return true;
            }
        }

        // Only blank rows remain: the sheet ends before this one.
        _ended = true;
        return false;
    }

    private bool ReadReaderRow()
    {
        if (!_reader.Read())
        {
            return false;
        }

        _readerRowIndex++;
        return true;
    }

    // Whether the row the reader is on has no value in any column.
    private bool IsReaderRowBlank()
    {
        for (int i = 0; i < _reader.FieldCount; i++)
        {
            var value = _reader.GetValue(i);
            if (value is not null && !(value is string stringValue && string.IsNullOrEmpty(stringValue)))
            {
                return false;
            }
        }

        return true;
    }
}

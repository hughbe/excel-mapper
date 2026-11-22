using ExcelDataReader;

namespace ExcelMapper.Readers;

/// <summary>
/// Reads the values of a cell split by a <see cref="CellSplitter"/>.
/// </summary>
internal sealed class SplitCellsReader(int column, CellSplitter splitter) : ICellsReader
{
    private string[]? _values;
    private ReadCellResult _readResult;
    private int _currentIndex;
    private int _position;
    // Whether there are more values to return when splitting without allocating an array.
    private bool _hasMore;

    public bool Start(IExcelDataReader reader, bool preserveFormatting, out int count)
    {
        _currentIndex = -1;
        _values = null;
        _position = 0;

        if (column >= reader.FieldCount)
        {
            count = 0;
            return false;
        }

        var readResult = new ReadCellResult(column, reader, preserveFormatting);

        _readResult = readResult;

        // An empty cell contains no values.
        var stringValue = readResult.GetString();
        if (string.IsNullOrEmpty(stringValue))
        {
            count = 0;
            _hasMore = false;
            return true;
        }

        _hasMore = true;

        // Try to avoid splitting if possible
        int directCount = splitter.GetCount(stringValue!);
        if (directCount >= 0)
        {
            count = directCount;
            return true;
        }

        // Fall back to splitting
        _values = splitter.GetValues(stringValue!);
        count = _values.Length;
        return true;
    }

    public bool TryGetNext(out ReadCellResult result)
    {
        _currentIndex++;
        if (_values == null)
        {
            // Keep going until the last segment has been returned. This includes an empty
            // final segment after a trailing separator, matching string.Split and the count
            // returned from Start.
            if (_hasMore)
            {
                var stringValue = _readResult.GetString()!;
                var remaining = stringValue.AsSpan(_position);
                var (advance, valueStart, valueLength) = splitter.GetNextValue(remaining);
                var value = stringValue.Substring(_position + valueStart, valueLength);
                if (advance >= 0)
                {
                    _position += advance;
                }
                else
                {
                    _hasMore = false;
                }

                result = new ReadCellResult(_readResult.ColumnIndex, value, _readResult.PreserveFormatting);
                return true;
            }
        }
        else
        {
            if (_currentIndex < _values.Length)
            {
                result = new ReadCellResult(_readResult.ColumnIndex, _values[_currentIndex], _readResult.PreserveFormatting);
                return true;
            }
        }

        result = default;
        return false;
    }

    public void Reset()
    {
        _currentIndex = -1;
        _values = null;
        _position = 0;
        _hasMore = false;
    }
}

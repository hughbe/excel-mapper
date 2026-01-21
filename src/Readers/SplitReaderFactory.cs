using System.Diagnostics.CodeAnalysis;
using ExcelDataReader;

namespace ExcelMapper.Readers;

/// <summary>
/// Reads the value of a cell and produces multiple values by splitting the string value
/// using the given separators.
/// </summary>
public abstract class SplitReaderFactory : ICellsReaderFactory, IColumnIndicesProviderCellReaderFactory, IColumnNamesProviderCellReaderFactory
{
    /// <summary>
    /// Gets or sets the options used to split the string value of the cell.
    /// </summary>
    public StringSplitOptions Options { get; set; }

    private ICellReaderFactory _readerFactory;

    /// <summary>
    /// Gets or sets the ICellReader that reads the string value of the cell
    /// before it is split.
    /// </summary>
    public ICellReaderFactory ReaderFactory
    {
        get => _readerFactory;
        set
        {
            ThrowHelpers.ThrowIfNull(value, nameof(value));
            _readerFactory = value;
        }
    }

    /// <summary>
    /// Constructs a reader that reads the string value of a cell and produces multiple values
    /// by splitting it.
    /// </summary>
    /// <param name="readerFactory">The ICellReaderFactory that reads the string value of the cell before it is split.</param>
    public SplitReaderFactory(ICellReaderFactory readerFactory)
    {
        ThrowHelpers.ThrowIfNull(readerFactory, nameof(readerFactory));
        _readerFactory = readerFactory;
    }

    /// <summary>
    /// Splits the given string value into multiple values.
    /// </summary>
    /// <param name="value">The string value to split.</param>
    /// <returns>The array of string values produced by splitting the string value.</returns>
    protected abstract string[] GetValues(string value);

    /// <summary>
    /// Gets the number of values that would be produced by splitting the given string value
    /// without allocating the intermediate array. Returns -1 if the split is required.
    /// </summary>
    /// <param name="value">The string value to analyze.</param>
    /// <returns>The count of values, or -1 if splitting is required.</returns>
    protected abstract int GetCount(string value);

    /// <summary>
    /// Gets the next value from the remaining string.
    /// Only called if GetCount returned a valid count.
    /// </summary>
    /// <param name="remaining">The remaining string to search.</param>
    /// <returns>A tuple with the number of characters to advance (value + separator), or -1 if this was the last value; the start offset of the value (after trimming leading whitespace); and the length of the value.</returns>
    protected abstract (int Advance, int ValueStart, int ValueLength) GetNextValue(ReadOnlySpan<char> remaining);

    /// <inheritdoc/>
    public ICellsReader? GetCellsReader(ExcelSheet sheet)
    {
        ICellReader? reader = _readerFactory.GetCellReader(sheet);
        if (reader == null)
        {
            return null;
        }

        return new Reader(reader, this);
    }

    /// <inheritdoc/>
    public IReadOnlyList<string>? GetColumnNames(ExcelSheet sheet)
    {
        if (_readerFactory is IColumnNameProviderCellReaderFactory nameProvider)
        {
            var columnName = nameProvider.GetColumnName(sheet);
            if (columnName != null && columnName != string.Empty)
            {
                return [columnName];
            }
        }

        return null;
    }

    /// <inheritdoc/>
    public IReadOnlyList<int>? GetColumnIndices(ExcelSheet sheet)
    {
        if (_readerFactory is IColumnIndexProviderCellReaderFactory indexProvider)
        {
            var columnIndex = indexProvider.GetColumnIndex(sheet);
            if (columnIndex != null && columnIndex != -1)
            {
                return [columnIndex.Value];
            }
        }

        return null;
    }

    private struct Reader(ICellReader Reader, SplitReaderFactory Splitter) : ICellsReader
    {
        private string[]? _values;
        private ReadCellResult _readResult;
        private int _currentIndex;
        private int _position;

        public bool Start(IExcelDataReader reader, bool preserveFormatting, out int count)
        {
            _currentIndex = -1;
            _values = null;
            _position = 0;

            if (!Reader.TryGetValue(reader, preserveFormatting, out var readResult))
            {
                count = 0;
                return false;
            }

            _readResult = readResult;

            var stringValue = readResult.GetString();
            if (stringValue == null)
            {
                count = 0;
                return true;
            }

            // Try to avoid splitting if possible
            int directCount = Splitter.GetCount(stringValue);
            if (directCount >= 0)
            {
                count = directCount;
                return true;
            }

            // Fall back to splitting
            _values = Splitter.GetValues(stringValue);
            count = _values.Length;
            return true;
        }

        public bool TryGetNext([NotNullWhen(true)] out ReadCellResult result)
        {
            _currentIndex++;
            if (_values == null)
            {
                if (!string.IsNullOrEmpty(_readResult.StringValue) && _position < _readResult.StringValue!.Length)
                {
                    var remaining = _readResult.StringValue.AsSpan(_position);
                    var (advance, valueStart, valueLength) = Splitter.GetNextValue(remaining);
                    var value = _readResult.StringValue!.Substring(_position + valueStart, valueLength);
                    _position += advance >= 0 ? advance : remaining.Length;

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
        }
    }
}

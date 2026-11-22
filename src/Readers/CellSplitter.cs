using System.Diagnostics.CodeAnalysis;
using ExcelDataReader;

namespace ExcelMapper.Readers;

/// <summary>
/// How the text of a cell is split into the values of a collection (see <see cref="SplitColumns"/>).
/// </summary>
internal abstract class CellSplitter
{
    public StringSplitOptions Options { get; set; }

    /// <summary>
    /// Splits the text into an array, for separators that can't be split without one.
    /// </summary>
    public abstract string[] GetValues(string value);

    /// <summary>
    /// Counts the values without splitting the text, or returns -1 if it can't be.
    /// </summary>
    public abstract int GetCount(string value);

    /// <summary>
    /// Finds the next value, when the values were counted without splitting.
    /// </summary>
    public abstract (int Advance, int ValueStart, int ValueLength) GetNextValue(ReadOnlySpan<char> remaining);
}

/// <summary>
/// Splits text at characters, such as ',' (the default).
/// </summary>
internal sealed class CharSplitter : CellSplitter
{
    private char[] _separators = [','];

    /// <summary>
    /// Gets or sets the separators used to split the string value of the cell.
    /// </summary>
    public char[] Separators
    {
        get => _separators;
        set
        {
            SeparatorUtilities.ValidateSeparators(value);
            _separators = value;
        }
    }

    /// <inheritdoc/>
    public override string[] GetValues(string value) => value.Split(Separators, Options);

    /// <inheritdoc/>
    public override int GetCount(string value)
    {
        // Can't easily calculate count with RemoveEmptyEntries without actually checking each segment
        if (Options.HasFlag(StringSplitOptions.RemoveEmptyEntries) || Separators.Length != 1)
        {
            return -1;
        }

#if NET8_0_OR_GREATER
        return value.AsSpan().Count(Separators[0]) + 1;
#else
        int count = 0;
        ReadOnlySpan<char> span = value.AsSpan();
        int index = 0;
        while ((index = span.IndexOf(Separators[0])) >= 0)
        {
            count++;
            span = span.Slice(index + 1);
        }

        return count + 1;
#endif
    }

    /// <inheritdoc/>
    public override (int Advance, int ValueStart, int ValueLength) GetNextValue(ReadOnlySpan<char> remaining)
    {
        var separator = Separators[0];
#if NET5_0_OR_GREATER
        var trimEntries = Options.HasFlag(StringSplitOptions.TrimEntries);
#else
        bool trimEntries = Options.HasFlag((StringSplitOptions)2);
#endif

        while (true)
        {
            // Get the index of the next separator.
            int separatorIndex = remaining.IndexOf(separator);
            
            if (separatorIndex >= 0)
            {
                ReadOnlySpan<char> value = remaining.Slice(0, separatorIndex);
                int valueStart = 0;
                int valueLength = separatorIndex;
                
                if (trimEntries)
                {
                    ReadOnlySpan<char> trimmed = value.Trim();
                    valueStart = value.Length - value.TrimStart().Length; // Leading whitespace offset
                    valueLength = trimmed.Length;
                }

                return (separatorIndex + 1, valueStart, valueLength);
            }

            // Last segment - no more separators.
            ReadOnlySpan<char> lastValue = remaining;
            int lastValueStart = 0;
            int lastValueLength = remaining.Length;
            
            if (trimEntries)
            {
                ReadOnlySpan<char> trimmed = lastValue.Trim();
                lastValueStart = lastValue.Length - lastValue.TrimStart().Length;
                lastValueLength = trimmed.Length;
            }

            return (-1, lastValueStart, lastValueLength);
        }
    }
}

/// <summary>
/// Splits text at strings, such as ", ".
/// </summary>
internal sealed class StringSplitter : CellSplitter
{
    private string[] _separators = [","];

    /// <summary>
    /// Gets or sets the separators used to split the string value of the cell.
    /// </summary>
    public string[] Separators
    {
        get => _separators;
        set
        {
            SeparatorUtilities.ValidateSeparators(value);
            _separators = value;
        }
    }

    /// <inheritdoc/>
    public override string[] GetValues(string value) => value.Split(Separators, Options);

    /// <inheritdoc/>
    public override int GetCount(string value)
    {
        // Can't easily calculate count with RemoveEmptyEntries without actually checking each segment
        if (Options.HasFlag(StringSplitOptions.RemoveEmptyEntries) || Separators.Length != 1)
        {
            return -1;
        }

        var separator = Separators[0];
#if NET8_0_OR_GREATER
        return value.AsSpan().Count(separator.AsSpan()) + 1;
#else
        int count = 0;
        ReadOnlySpan<char> span = value.AsSpan();
        int index = 0;
        while ((index = span.IndexOf(separator.AsSpan())) >= 0)
        {
            count++;
            span = span.Slice(index + separator.Length);
        }

        return count + 1;
#endif
    }

    /// <inheritdoc/>
    public override (int Advance, int ValueStart, int ValueLength) GetNextValue(ReadOnlySpan<char> remaining)
    {
        var separator = Separators[0];
#if NET5_0_OR_GREATER
        var trimEntries = Options.HasFlag(StringSplitOptions.TrimEntries);
#else
        var trimEntries = Options.HasFlag((StringSplitOptions)2);
#endif

        while (true)
        {
            // Get the index of the next separator.
            var separatorIndex = remaining.IndexOf(separator.AsSpan());
            
            if (separatorIndex >= 0)
            {
                ReadOnlySpan<char> value = remaining.Slice(0, separatorIndex);
                int valueStart = 0;
                int valueLength = separatorIndex;
                
                if (trimEntries)
                {
                    ReadOnlySpan<char> trimmed = value.Trim();
                    valueStart = value.Length - value.TrimStart().Length; // Leading whitespace offset
                    valueLength = trimmed.Length;
                }

                return (separatorIndex + separator.Length, valueStart, valueLength);
            }

            // Last segment - no more separators.
            ReadOnlySpan<char> lastValue = remaining;
            int lastValueStart = 0;
            int lastValueLength = remaining.Length;
            
            if (trimEntries)
            {
                ReadOnlySpan<char> trimmed = lastValue.Trim();
                lastValueStart = lastValue.Length - lastValue.TrimStart().Length;
                lastValueLength = trimmed.Length;
            }

            return (-1, lastValueStart, lastValueLength);
        }
    }
}

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

    public bool TryGetNext([NotNullWhen(true)] out ReadCellResult result)
    {
        _currentIndex++;
        if (_values == null)
        {
            // Keep going until the last segment has been returned. This includes an empty
            // final segment after a trailing separator, matching string.Split and the count
            // returned from Start.
            if (_hasMore)
            {
                var stringValue = _readResult.StringValue!;
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

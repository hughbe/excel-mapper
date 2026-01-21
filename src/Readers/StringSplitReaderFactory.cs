namespace ExcelMapper.Readers;

/// <summary>
/// Reads the value of a cell and produces multiple values by splitting the string value
/// using the given separators.
/// </summary>
public class StringSplitReaderFactory : SplitReaderFactory
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

    /// <summary>
    /// Constructs a reader that reads the string value of a cell and produces multiple values
    /// by splitting it.
    /// </summary>
    /// <param name="cellReader">The ICellReaderFactory that reads the string value of the cell before it is split.</param>
    public StringSplitReaderFactory(ICellReaderFactory cellReader) : base(cellReader)
    {
    }

    /// <inheritdoc/>
    protected override string[] GetValues(string value) => value.Split(Separators, Options);

    /// <inheritdoc/>
    protected override int GetCount(string value)
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
    protected override (int Advance, int ValueStart, int ValueLength) GetNextValue(ReadOnlySpan<char> remaining)
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


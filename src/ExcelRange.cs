namespace ExcelMapper;

/// <summary>
/// Represents a range of rows and columns in an Excel worksheet.
/// </summary>
/// <remarks>
/// The default value of this struct (for example <c>default(ExcelRange)</c>) represents every
/// row and column in the worksheet, the same as <see cref="ExcelRange()"/>.
/// </remarks>
public readonly struct ExcelRange : IEquatable<ExcelRange>
{
    /// <summary>
    /// The maximum number of rows in an Excel worksheet (row 1048576).
    /// </summary>
    internal const int MaxRows = 1_048_576;

    /// <summary>
    /// The maximum number of columns in an Excel worksheet (column XFD).
    /// </summary>
    internal const int MaxColumns = 16_384;

    // Nullable so that default(ExcelRange) means "all rows and columns" rather than "0..0".
    private readonly Range? _rows;
    private readonly Range? _columns;

    /// <summary>
    /// Gets the row range.
    /// </summary>
    public Range Rows => _rows ?? Range.All;

    /// <summary>
    /// Gets the column range.
    /// </summary>
    public Range Columns => _columns ?? Range.All;

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelRange"/> struct representing every row
    /// and column in the worksheet.
    /// </summary>
    public ExcelRange()
    {
        _rows = Range.All;
        _columns = Range.All;
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelRange"/> struct.
    /// </summary>
    /// <param name="rowStart">The inclusive start index of the row range.</param>
    /// <param name="rowEnd">The inclusive end index of the row range.</param>
    /// <param name="columnStart">The inclusive start index of the column range.</param>
    /// <param name="columnEnd">The inclusive end index of the column range.</param>
    public ExcelRange(int rowStart, int? rowEnd, int columnStart, int? columnEnd)
    {
        ThrowHelpers.ThrowIfNegative(rowStart, nameof(rowStart));
        if (rowEnd.HasValue)
        {
            ThrowHelpers.ThrowIfNegative(rowEnd.Value, nameof(rowEnd));
            ThrowHelpers.ThrowIfGreaterThan(rowStart, rowEnd.Value, nameof(rowStart));
        }
        ThrowHelpers.ThrowIfNegative(columnStart, nameof(columnStart));
        if (columnEnd.HasValue)
        {
            ThrowHelpers.ThrowIfNegative(columnEnd.Value, nameof(columnEnd));
            ThrowHelpers.ThrowIfGreaterThan(columnStart, columnEnd.Value, nameof(columnStart));
        }

        _rows = rowEnd.HasValue ? rowStart..(rowEnd.Value + 1) : rowStart..;
        _columns = columnEnd.HasValue ? columnStart..(columnEnd.Value + 1) : columnStart..;
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelRange"/> struct.
    /// </summary>
    /// <param name="rows">The row range.</param>
    /// <param name="columns">The column range.</param>
    public ExcelRange(Range rows, Range columns)
    {
        _rows = rows;
        _columns = columns;
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelRange"/> struct from an Excel address string.
    /// </summary>
    /// <param name="address">The Excel address string (e.g., "A1:B10", "C5", "$A$1", "A:C", "1:5").</param>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="address"/> is null.</exception>
    /// <exception cref="ArgumentException">Thrown when <paramref name="address"/> is empty or has an invalid format.</exception>
    public ExcelRange(string address)
    {
        this = Parse(address);
    }

    /// <summary>
    /// Parses an Excel address string such as "A1", "A1:B10", "$A$1:$B$10", "A:C" or "1:5".
    /// Column letters are case-insensitive. Rows must be between 1 and 1048576 and columns between
    /// A and XFD, the limits of an Excel worksheet.
    /// </summary>
    /// <param name="address">The Excel address string.</param>
    /// <returns>The range represented by the address.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="address"/> is null.</exception>
    /// <exception cref="ArgumentException">Thrown when <paramref name="address"/> is empty or has an invalid format.</exception>
    public static ExcelRange Parse(string address)
    {
        ThrowHelpers.ThrowIfNull(address, nameof(address));

        if (!TryParseCore(address.AsSpan(), out var result, out var error))
        {
            throw new ArgumentException($"Invalid Excel address '{address}': {error} Expected format like 'A1', 'A1:B10', 'A:C', or '1:5'.", nameof(address));
        }

        return result;
    }

    /// <summary>
    /// Tries to parse an Excel address string such as "A1", "A1:B10", "$A$1:$B$10", "A:C" or "1:5".
    /// </summary>
    /// <param name="address">The Excel address string.</param>
    /// <param name="result">The range represented by the address, or a range of every row and
    /// column if the address is invalid.</param>
    /// <returns>True if the address was parsed, else false.</returns>
    public static bool TryParse(string? address, out ExcelRange result)
    {
        if (address is null)
        {
            result = default;
            return false;
        }

        return TryParseCore(address.AsSpan(), out result, out _);
    }

    private static bool TryParseCore(ReadOnlySpan<char> address, out ExcelRange result, out string? error)
    {
        result = default;
        var span = address.Trim();
        if (span.IsEmpty)
        {
            error = "The address is empty.";
            return false;
        }

        int colonIndex = span.IndexOf(':');
        if (colonIndex == -1)
        {
            // Single cell reference like "A1".
            if (!TryParseCellReference(span, out int row, out int column, out error))
            {
                return false;
            }

            result = new ExcelRange(row..(row + 1), column..(column + 1));
            return true;
        }

        var start = span[..colonIndex].Trim();
        var end = span[(colonIndex + 1)..].Trim();
        if (end.IndexOf(':') != -1)
        {
            error = "The address contains multiple colons.";
            return false;
        }

        // Column-only ranges like "A:C".
        if (IsColumnOnlyReference(start))
        {
            if (!TryParseColumnLetters(start, out int startColumn) || !TryParseColumnLetters(end, out int endColumn))
            {
                error = "Columns must be letters between A and XFD.";
                return false;
            }

            if (startColumn > endColumn)
            {
                error = $"Start column '{start.ToString()}' is after end column '{end.ToString()}'.";
                return false;
            }

            result = new ExcelRange(Range.All, startColumn..(endColumn + 1));
            error = null;
            return true;
        }

        // Row-only ranges like "1:5".
        if (IsRowOnlyReference(start))
        {
            if (!TryParseRowNumber(start, out int startRow) || !TryParseRowNumber(end, out int endRow))
            {
                error = $"Rows must be numbers between 1 and {MaxRows}.";
                return false;
            }

            if (startRow > endRow)
            {
                error = $"Start row '{start.ToString()}' is after end row '{end.ToString()}'.";
                return false;
            }

            result = new ExcelRange(startRow..(endRow + 1), Range.All);
            error = null;
            return true;
        }

        // Cell ranges like "A1:B10".
        if (!TryParseCellReference(start, out int startCellRow, out int startCellColumn, out error) ||
            !TryParseCellReference(end, out int endCellRow, out int endCellColumn, out error))
        {
            return false;
        }

        if (startCellRow > endCellRow)
        {
            error = $"Start row {startCellRow + 1} is after end row {endCellRow + 1}.";
            return false;
        }

        if (startCellColumn > endCellColumn)
        {
            error = $"Start column {GetColumnLetters(startCellColumn)} is after end column {GetColumnLetters(endCellColumn)}.";
            return false;
        }

        result = new ExcelRange(startCellRow..(endCellRow + 1), startCellColumn..(endCellColumn + 1));
        error = null;
        return true;
    }

    private static ReadOnlySpan<char> TrimAbsoluteMarker(ReadOnlySpan<char> reference)
        => !reference.IsEmpty && reference[0] == '$' ? reference[1..] : reference;

    private static bool IsAsciiLetter(char c) => (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z');

    private static bool IsAsciiDigit(char c) => c >= '0' && c <= '9';

    private static bool IsRowOnlyReference(ReadOnlySpan<char> reference)
    {
        reference = TrimAbsoluteMarker(reference);
        if (reference.IsEmpty)
        {
            return false;
        }

        foreach (var c in reference)
        {
            if (!IsAsciiDigit(c))
            {
                return false;
            }
        }

        return true;
    }

    private static bool IsColumnOnlyReference(ReadOnlySpan<char> reference)
    {
        reference = TrimAbsoluteMarker(reference);
        if (reference.IsEmpty)
        {
            return false;
        }

        foreach (var c in reference)
        {
            if (!IsAsciiLetter(c))
            {
                return false;
            }
        }

        return true;
    }

    /// <summary>
    /// Parses a one-based row number such as "5" or "$5" into a zero-based row index.
    /// </summary>
    private static bool TryParseRowNumber(ReadOnlySpan<char> rowString, out int row)
    {
        rowString = TrimAbsoluteMarker(rowString);
        row = -1;
        if (rowString.IsEmpty)
        {
            return false;
        }

        long rowNumber = 0;
        foreach (var c in rowString)
        {
            if (!IsAsciiDigit(c))
            {
                return false;
            }

            rowNumber = rowNumber * 10 + (c - '0');
            if (rowNumber > MaxRows)
            {
                return false;
            }
        }

        if (rowNumber < 1)
        {
            return false;
        }

        // Convert from 1-based Excel row to 0-based index.
        row = (int)rowNumber - 1;
        return true;
    }

    /// <summary>
    /// Parses column letters such as "A", "ab" or "$XFD" into a zero-based column index.
    /// </summary>
    private static bool TryParseColumnLetters(ReadOnlySpan<char> columnString, out int column)
    {
        columnString = TrimAbsoluteMarker(columnString);
        column = -1;
        if (columnString.IsEmpty)
        {
            return false;
        }

        int columnNumber = 0;
        foreach (var c in columnString)
        {
            if (!IsAsciiLetter(c))
            {
                return false;
            }

            columnNumber = columnNumber * 26 + (char.ToUpperInvariant(c) - 'A' + 1);
            if (columnNumber > MaxColumns)
            {
                return false;
            }
        }

        // Convert from 1-based Excel column to 0-based index.
        column = columnNumber - 1;
        return true;
    }

    /// <summary>
    /// Parses a cell reference such as "B5", "$B$5", "$B5" or "B$5".
    /// </summary>
    private static bool TryParseCellReference(ReadOnlySpan<char> cellReference, out int row, out int column, out string? error)
    {
        row = -1;
        column = -1;

        // Split into the column letters and the row number, each optionally preceded by "$".
        int index = !cellReference.IsEmpty && cellReference[0] == '$' ? 1 : 0;
        int lettersStart = index;
        while (index < cellReference.Length && IsAsciiLetter(cellReference[index]))
        {
            index++;
        }

        var columnPart = cellReference[lettersStart..index];
        var rowPart = cellReference[index..];
        if (columnPart.IsEmpty || TrimAbsoluteMarker(rowPart).IsEmpty)
        {
            error = $"Invalid cell reference '{cellReference.ToString()}'.";
            return false;
        }

        if (!TryParseColumnLetters(columnPart, out column))
        {
            error = $"Invalid column '{columnPart.ToString()}'. Columns must be letters between A and XFD.";
            return false;
        }

        if (!TryParseRowNumber(rowPart, out row))
        {
            error = $"Invalid row '{rowPart.ToString()}'. Rows must be numbers between 1 and {MaxRows}.";
            return false;
        }

        error = null;
        return true;
    }

    /// <summary>
    /// Gets the Excel column letters (for example "A", "Z" or "AA") of a zero-based column index.
    /// </summary>
    internal static string GetColumnLetters(int columnIndex)
    {
        ThrowHelpers.ThrowIfNegative(columnIndex, nameof(columnIndex));

        // Enough for any int: 26^7 > int.MaxValue.
        Span<char> buffer = stackalloc char[7];
        int position = buffer.Length;
        long number = (long)columnIndex + 1;
        while (number > 0)
        {
            number--;
            buffer[--position] = (char)('A' + (int)(number % 26));
            number /= 26;
        }

        return buffer[position..].ToString();
    }

    /// <summary>
    /// Gets the Excel address (for example "B11") of a cell given its zero-based row and column indices.
    /// </summary>
    internal static string GetCellAddress(int rowIndex, int columnIndex)
    {
        ThrowHelpers.ThrowIfNegative(rowIndex, nameof(rowIndex));
        return $"{GetColumnLetters(columnIndex)}{(long)rowIndex + 1}";
    }

    private static bool TryGetBounds(Range range, out int start, out int end)
    {
        start = range.Start.Value;
        end = range.End.Value;
        return !range.Start.IsFromEnd && !range.End.IsFromEnd && end > start;
    }

    /// <summary>
    /// Returns the Excel address of the range (for example "A1:B10", "C5", "A:C" or "1:5") if it can
    /// be expressed as one, otherwise a description of the row and column ranges.
    /// </summary>
    public override string ToString()
    {
        var rowsAll = Rows.Equals(Range.All);
        var columnsAll = Columns.Equals(Range.All);
        var rowsBounded = TryGetBounds(Rows, out var rowStart, out var rowEnd);
        var columnsBounded = TryGetBounds(Columns, out var columnStart, out var columnEnd);

        if (rowsBounded && columnsBounded)
        {
            var startAddress = GetCellAddress(rowStart, columnStart);
            if (rowEnd == rowStart + 1 && columnEnd == columnStart + 1)
            {
                return startAddress;
            }

            return $"{startAddress}:{GetCellAddress(rowEnd - 1, columnEnd - 1)}";
        }

        if (rowsAll && columnsBounded)
        {
            return $"{GetColumnLetters(columnStart)}:{GetColumnLetters(columnEnd - 1)}";
        }

        if (rowsBounded && columnsAll)
        {
            return $"{(long)rowStart + 1}:{rowEnd}";
        }

        return $"Rows: {Rows}, Columns: {Columns}";
    }

    /// <inheritdoc/>
    public bool Equals(ExcelRange other) => Rows.Equals(other.Rows) && Columns.Equals(other.Columns);

    /// <inheritdoc/>
    public override bool Equals(object? obj) => obj is ExcelRange other && Equals(other);

    /// <inheritdoc/>
    public override int GetHashCode()
    {
        unchecked
        {
            return (Rows.GetHashCode() * 397) ^ Columns.GetHashCode();
        }
    }

    /// <summary>
    /// Returns whether two ranges contain the same rows and columns.
    /// </summary>
    public static bool operator ==(ExcelRange left, ExcelRange right) => left.Equals(right);

    /// <summary>
    /// Returns whether two ranges contain different rows or columns.
    /// </summary>
    public static bool operator !=(ExcelRange left, ExcelRange right) => !left.Equals(right);
}

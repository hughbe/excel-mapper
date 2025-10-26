using System.Linq;

namespace ExcelMapper;

/// <summary>
/// Represents a range of rows and columns in an Excel worksheet.
/// </summary>
public struct ExcelRange
{
    /// <summary>
    /// Gets the row range.
    /// </summary>
    public Range Rows { get; }

    /// <summary>
    /// Gets the column range.
    /// </summary>
    public Range Columns { get; }

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelRange"/> struct.
    /// </summary>
    public ExcelRange()
    {
        Rows = Range.All;
        Columns = Range.All;
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
        ArgumentOutOfRangeException.ThrowIfNegative(rowStart);
        if (rowEnd.HasValue)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(rowEnd.Value, nameof(rowEnd));
            ArgumentOutOfRangeException.ThrowIfGreaterThan(rowStart, rowEnd.Value, nameof(rowStart));
        }
        ArgumentOutOfRangeException.ThrowIfNegative(columnStart);
        if (columnEnd.HasValue)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(columnEnd.Value, nameof(columnEnd));
            ArgumentOutOfRangeException.ThrowIfGreaterThan(columnStart, columnEnd.Value, nameof(columnStart));
        }

        Rows = rowEnd.HasValue ? rowStart..(rowEnd.Value + 1) : rowStart..;
        Columns = columnEnd.HasValue ? columnStart..(columnEnd.Value + 1) : columnStart..;
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelRange"/> struct.
    /// </summary>
    /// <param name="rows">The row range.</param>
    /// <param name="columns">The column range.</param>
    public ExcelRange(Range rows, Range columns)
    {
        Rows = rows;
        Columns = columns;
    }
    
    public ExcelRange(string address)
    {
        this = Parse(address);
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelRange"/> struct from an Excel address string.
    /// </summary>
    /// <param name="address">The Excel address string (e.g., "A1:B10", "C5", "A1:A", "1:5").</param>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="address"/> is null.</exception>
    /// <exception cref="ArgumentException">Thrown when <paramref name="address"/> is empty or has an invalid format.</exception>
    public static ExcelRange Parse(string address)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(address);

        var span = address.AsSpan().Trim();
        int colonIndex = span.IndexOf(':');

        if (colonIndex == -1)
        {
            // Single cell reference like "A1"
            ParseCellReference(span, out int row, out int column);
            return new ExcelRange(
                rows: row..(row + 1),
                columns: column..(column + 1)
            );
        }
        else
        {
            // Range reference like "A1:B10"
            var start = span[..colonIndex].Trim();
            var end = span[(colonIndex + 1)..].Trim();

            // Check for multiple colons
            if (end.IndexOf(':') != -1)
            {
                throw new ArgumentException($"Invalid Excel address format: '{address}'. Expected format like 'A1', 'A1:B10', 'A:C', or '1:5'.", nameof(address));
            }

            var startIsRowOnly = IsRowOnlyReference(start);
            var startIsColumnOnly = IsColumnOnlyReference(start);

            // Handle column-only ranges like "A:C"
            if (startIsColumnOnly)
            {
                int startColumn = ParseColumnLetters(start);
                int endColumn = ParseColumnLetters(end);
                
                if (startColumn > endColumn)
                {
                    throw new ArgumentException($"Invalid Excel address: start column '{start}' is after end column '{end}'.", nameof(address));
                }

                return new ExcelRange(
                    rows: Range.All,
                    columns: startColumn..(endColumn + 1)
                );
            }
            // Handle row-only ranges like "1:5"
            else if (startIsRowOnly)
            {
                int startRow = ParseRowNumber(start);
                int endRow = ParseRowNumber(end);
                
                if (startRow > endRow)
                {
                    throw new ArgumentException($"Invalid Excel address: start row '{start}' is after end row '{end}'.", nameof(address));
                }

                return new ExcelRange(
                    rows: startRow..(endRow + 1),
                    columns: Range.All
                );
            }
            // Handle full cell ranges like "A1:B10"
            else
            {
                ParseCellReference(start, out int startRow, out int startColumn);
                ParseCellReference(end, out int endRow, out int endColumn);

                if (startRow > endRow)
                {
                    throw new ArgumentException($"Invalid Excel address: start row {startRow + 1} is after end row {endRow + 1}.", nameof(address));
                }

                if (startColumn > endColumn)
                {
                    throw new ArgumentException($"Invalid Excel address: start column is after end column.", nameof(address));
                }

                return new ExcelRange(
                    rows: startRow..(endRow + 1),
                    columns: startColumn..(endColumn + 1)
                );
            }
        }
    }

    private static bool IsRowOnlyReference(ReadOnlySpan<char> reference)
    {
        if (reference.IsEmpty)
        {
            return false;
        }

        foreach (var c in reference)
        {
            if (!char.IsDigit(c))
            {
                return false;
            }
        }

        return true;
    }

    private static bool IsColumnOnlyReference(ReadOnlySpan<char> reference)
    {
        if (reference.IsEmpty)
        {
            return false;
        }

        foreach (var c in reference)
        {
            if (!char.IsLetter(c))
            {
                return false;
            }
        }

        return true;
    }

    /// <summary>
    /// Tries to parse an Excel address string into an <see cref="ExcelRange"/>.
    /// </summary>
    /// <param name="address">The Excel address string.</param>
    public static bool TryParse(string address, out ExcelRange result)
    {
        try
        {
            result = Parse(address);
            return true;
        }
        catch
        {
            result = default;
            return false;
        }
    }

    private static int ParseRowNumber(ReadOnlySpan<char> rowString)
    {
#if NET8_0_OR_GREATER
        if (!int.TryParse(rowString, out int rowNumber) || rowNumber < 1)
#else
        if (!int.TryParse(rowString.ToString(), out int rowNumber) || rowNumber < 1)
#endif
        {
            throw new ArgumentException($"Invalid row number: '{rowString}'. Row numbers must be positive integers.", "address");
        }

        // Convert from 1-based Excel row to 0-based index
        return rowNumber - 1;
    }

    private static int ParseColumnLetters(ReadOnlySpan<char> columnString)
    {
        if (columnString.IsEmpty)
        {
            throw new ArgumentException($"Invalid column letters: column must contain only letters.", "address");
        }

        int column = 0;
        foreach (var c in columnString)
        {
            char upper = char.ToUpperInvariant(c);
            if (!char.IsLetter(upper))
            {
                throw new ArgumentException($"Invalid column letters: '{columnString}'. Column must contain only letters.", "address");
            }
            column = column * 26 + (upper - 'A' + 1);
        }

        // Convert from 1-based Excel column to 0-based index
        return column - 1;
    }

    private static void ParseCellReference(ReadOnlySpan<char> cellReference, out int row, out int column)
    {
        if (cellReference.IsEmpty)
        {
            throw new ArgumentException("Cell reference cannot be empty.", "address");
        }

        // Split into letters and numbers
        int firstDigitIndex = -1;
        for (int i = 0; i < cellReference.Length; i++)
        {
            if (char.IsDigit(cellReference[i]))
            {
                firstDigitIndex = i;
                break;
            }
        }

        if (firstDigitIndex == -1 || firstDigitIndex == 0)
        {
            throw new ArgumentException($"Invalid cell reference: '{cellReference}'. Expected format like 'A1'.", "address");
        }

        ReadOnlySpan<char> columnPart = cellReference[..firstDigitIndex];
        ReadOnlySpan<char> rowPart = cellReference[firstDigitIndex..];

        column = ParseColumnLetters(columnPart);
        row = ParseRowNumber(rowPart);
    }
}

using System.Globalization;
using ExcelDataReader;
using ExcelNumberFormat;

namespace ExcelMapper.Abstractions;

/// <summary>
/// Metadata about the output of reading the value of a single cell.
/// </summary>
public struct ReadCellResult
{
    private object? _value;
    private string? _stringValue;

    /// <summary>
    /// The reader for the cell.
    /// </summary>
    public IExcelDataReader? Reader { get; }

    /// <summary>
    /// The index of the column that contains the cell.
    /// </summary>
    public int ColumnIndex { get; }

    /// <summary>
    /// Gets the string value of the cell.
    /// </summary>
    public string? GetString()
    {
        if (Reader == null || _stringValue != null)
        {
            return _stringValue;
        }

        var value = GetValue();
        if (PreserveFormatting)
        {
            var numberFormatString = Reader.GetNumberFormatString(ColumnIndex);
            var numberFormat = new NumberFormat(numberFormatString);
            _stringValue = numberFormat.Format(value, CultureInfo.CurrentCulture);
        }
        else
        {
            _stringValue = value?.ToString();
        }

        return _stringValue;
    }

    /// <summary>
    /// Gets the value of the cell.
    /// </summary>
    /// <returns>The value of the cell.</returns>
    public object? GetValue()
    {
        if (Reader == null || _value != null)
        {
            return _value;
        }

        var value = Reader.GetValue(ColumnIndex);
        // Cache the string value for performance to save us needing to read it again.
        if (value is string stringValue)
        {
            _stringValue = stringValue;
        }

        _value = value;
        return value;
    }

    /// <summary>
    /// Gets whether the cell is empty.
    /// </summary>
    public bool IsEmpty()
    {
        // If we preserve formatting, get the value as a string.
        if (PreserveFormatting)
        {
            return string.IsNullOrEmpty(GetString());
        }

        var value = GetValue();
        return value is null || (value is string stringValue && string.IsNullOrEmpty(stringValue));
    }

    /// <summary>
    /// Gets the string value of the cell, ignoring formatting
    /// </summary>
    public string? StringValue => GetString();

    /// <summary>
    /// Gets whether to preserve number formatting options when reading string values.
    /// </summary>
    public bool PreserveFormatting { get; }

    /// <summary>
    /// Constructs an object describing the output of reading the value of a single cell.
    /// </summary>
    /// <param name="columnIndex">The index of the column that contains the cell.</param>
    /// <param name="reader">The reader for the cell.</param>
    /// <param name="preserveFormatting">Whether or not to preserve formatting when reading string values.</param>
    public ReadCellResult(int columnIndex, IExcelDataReader reader, bool preserveFormatting)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(columnIndex);
        ArgumentNullException.ThrowIfNull(reader);

        ColumnIndex = columnIndex;
        Reader = reader;
        PreserveFormatting = preserveFormatting;
    }

    /// <summary>
    /// Constructs an object describing the output of reading the value of a single cell.
    /// </summary>
    /// <param name="columnIndex">The index of the column that contains the cell.</param>
    /// <param name="stringValue">The string value of the cell.</param>
    public ReadCellResult(int columnIndex, string? stringValue, bool preserveFormatting)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(columnIndex);

        ColumnIndex = columnIndex;
        _stringValue = stringValue;
        _value = stringValue;
        PreserveFormatting = preserveFormatting;
    }
}

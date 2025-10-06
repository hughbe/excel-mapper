using System.Globalization;
using ExcelDataReader;
using ExcelNumberFormat;

namespace ExcelMapper.Abstractions;

/// <summary>
/// Metadata about the output of reading the value of a single cell.
/// </summary>
public struct ReadCellResult
{
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
        if (Reader == null || (_stringValue != null && !PreserveFormatting))
        {
            return _stringValue;
        }

        if (PreserveFormatting)
        {
            var numberFormatString = Reader.GetNumberFormatString(ColumnIndex);
            var numberFormat = new NumberFormat(numberFormatString);
            return numberFormat.Format(Reader[ColumnIndex], CultureInfo.CurrentCulture);
        }
        else
        {
            _stringValue = Reader[ColumnIndex]?.ToString();
            return _stringValue;
        }
    }

    public object? GetValue()
    {
        if (Reader == null || _stringValue != null)
        {
            return _stringValue;
        }

        var value = Reader.GetValue(ColumnIndex);
        // Cache the string value for performance to save us needing to read it again.
        if (value is string stringValue)
        {
            _stringValue = stringValue;
        }

        return value;
    }

    public bool IsEmpty()
    {
        if (Reader == null || _stringValue != null)
        {
            return string.IsNullOrEmpty(_stringValue);
        }

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
        ColumnIndex = columnIndex;
        _stringValue = stringValue;
        PreserveFormatting = preserveFormatting;
    }
}

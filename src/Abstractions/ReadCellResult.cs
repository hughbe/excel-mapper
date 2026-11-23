using System.Globalization;
using ExcelDataReader;
using ExcelNumberFormat;

namespace ExcelMapper.Abstractions;

/// <summary>
/// Metadata about the output of reading the value of a single cell.
/// </summary>
internal struct ReadCellResult
{
    private object? _value;
    private string? _stringValue;

    /// <summary>
    /// The reader for the cell.
    /// </summary>
    public IExcelDataReader? Reader { get; }

    /// <summary>
    /// The culture to parse and format the cell's text with: the importer's
    /// <see cref="ExcelImporterConfiguration.Culture"/>, set when the cell is mapped
    /// (<see cref="CellPipeline{T}.Map"/>). Invariant if it isn't set.
    /// </summary>
    public IFormatProvider Culture
    {
        readonly get => _culture ?? CultureInfo.InvariantCulture;
        set => _culture = value;
    }

    private IFormatProvider? _culture;

    /// <summary>
    /// The index of the column that contains the cell.
    /// </summary>
    public int ColumnIndex { get; }

    /// <summary>
    /// Gets the string value of the cell.
    /// </summary>
    public string? GetString()
    {
        if (_stringValue != null)
        {
            return _stringValue;
        }

        if (Reader == null)
        {
            // A value created without a reader, such as a transformer's result.
            _stringValue = Format(_value);
            return _stringValue;
        }

        var value = GetValue();
        if (PreserveFormatting)
        {
            var numberFormatString = Reader.GetNumberFormatString(ColumnIndex);
            var numberFormat = new NumberFormat(numberFormatString);
            _stringValue = numberFormat.Format(value, Culture as CultureInfo ?? CultureInfo.InvariantCulture);
        }
        else
        {
            _stringValue = Format(value);
        }

        return _stringValue;
    }

    /// <summary>
    /// Gets the value of the cell.
    /// </summary>
    /// <returns>The value of the cell.</returns>
    // The text of a typed value, such as a number or date, in the cell's culture.
    private readonly string? Format(object? value)
        => value is IFormattable formattable ? formattable.ToString(null, Culture) : value?.ToString();

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
        ThrowHelpers.ThrowIfNegative(columnIndex, nameof(columnIndex));
        ThrowHelpers.ThrowIfNull(reader, nameof(reader));

        ColumnIndex = columnIndex;
        Reader = reader;
        PreserveFormatting = preserveFormatting;
    }

    /// <summary>
    /// Constructs an object describing the output of reading the value of a single cell.
    /// </summary>
    /// <param name="columnIndex">The index of the column that contains the cell.</param>
    /// <param name="value">The cell's value, such as the result of an <see cref="ExcelMapper.ICellTransformer"/>.</param>
    internal ReadCellResult(int columnIndex, object? value)
    {
        ColumnIndex = columnIndex;
        _value = value;
        _stringValue = value as string;
    }

    public ReadCellResult(int columnIndex, string? stringValue, bool preserveFormatting)
    {
        ThrowHelpers.ThrowIfNegative(columnIndex, nameof(columnIndex));

        ColumnIndex = columnIndex;
        _stringValue = stringValue;
        _value = stringValue;
        PreserveFormatting = preserveFormatting;
    }
}

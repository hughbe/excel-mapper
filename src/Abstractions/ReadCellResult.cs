using System.Collections.Concurrent;
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
        if (_isError)
        {
            return _stringValue;
        }

        if (PreserveFormatting)
        {
            var numberFormat = GetNumberFormat(Reader.GetNumberFormatString(ColumnIndex));
            _stringValue = numberFormat.Format(value, Culture as CultureInfo ?? CultureInfo.InvariantCulture);
        }
        else
        {
            _stringValue = Format(value);
        }

        return _stringValue;
    }

    // Parsed number formats by format string. Parsing one for every cell was most of the cost of
    // preserving formatting. A workbook has few formats, but one from an untrusted file could have
    // any number, so the cache is emptied when it's full rather than growing without limit.
    // NumberFormat isn't changed after it's parsed, so it can be shared between threads.
    private const int MaxCachedNumberFormats = 256;
    private static readonly ConcurrentDictionary<string, NumberFormat> s_numberFormats = new();
    private static NumberFormat? s_noNumberFormat;

    private static NumberFormat GetNumberFormat(string? formatString)
    {
        if (formatString is null)
        {
            // CSV files have no formats.
            return s_noNumberFormat ??= new NumberFormat(null);
        }

        if (s_numberFormats.TryGetValue(formatString, out var numberFormat))
        {
            return numberFormat;
        }

        if (s_numberFormats.Count >= MaxCachedNumberFormats)
        {
            s_numberFormats.Clear();
        }

        numberFormat = new NumberFormat(formatString);
        s_numberFormats[formatString] = numberFormat;
        return numberFormat;
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
        if (Reader == null || _value != null || _read)
        {
            return _value;
        }

        // Empty cells are read once too, as checking for an error is a second call.
        _read = true;
        var value = Reader.GetValue(ColumnIndex);
        // Cache the string value for performance to save us needing to read it again.
        if (value is string stringValue)
        {
            _stringValue = stringValue;
        }
        else if (value is null && GetErrorText(Reader.GetCellError(ColumnIndex)) is { } errorText)
        {
            // A formula error, such as #DIV/0!, which the reader gives as null: it isn't empty, and
            // its text is the error Excel shows.
            _isError = true;
            _stringValue = errorText;
            value = errorText;
        }

        _value = value;
        return value;
    }

    private bool _isError;
    private bool _read;

    /// <summary>
    /// Gets whether the cell holds a formula error, such as <c>#DIV/0!</c> or <c>#N/A</c>, whose
    /// text is the error.
    /// </summary>
    public bool IsError()
    {
        GetValue();
        return _isError;
    }

    // The text Excel shows for a formula error.
    internal static string? GetErrorText(CellError? error) => error switch
    {
        null => null,
        CellError.NULL => "#NULL!",
        CellError.DIV0 => "#DIV/0!",
        CellError.VALUE => "#VALUE!",
        CellError.REF => "#REF!",
        CellError.NAME => "#NAME?",
        CellError.NUM => "#NUM!",
        CellError.NA => "#N/A",
        CellError.GETTING_DATA => "#GETTING_DATA",
        _ => "#" + error,
    };

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

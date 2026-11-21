using System.Globalization;

namespace ExcelMapper.Mappers;

/// <summary>
/// A mapper that tries to map the value of a cell to a DateTime.
/// </summary>
public class DateTimeMapper : ICellMapper, IFormatsCellMapper, ITypedCellMapper<DateTime>, ITypedCellMapper<DateTime?>
{
    private string[] _formats = ["G"];

    /// <summary>
    /// Gets or sets the date formats used to map the value to a DateTime.
    /// This defaults to "G" - the default Excel format.
    /// While the formats haven't been set, values that don't match also try the culture's short date ("d") and short date and time ("g"), and ISO 8601, such as "2024-03-04" or "2024-03-04T10:30:00Z".
    /// </summary>
    private bool _formatsSet;

    public string[] Formats
    {
        get => _formats;
        set
        {
            FormatUtilities.ValidateFormats(value);
            _formats = value;
            _formatsSet = true;
        }
    }

    /// <summary>
    /// Gets or sets the IFormatProvider used to map the value to a DateTime.
    /// </summary>
    public IFormatProvider? Provider { get; set; }

    /// <summary>
    /// Gets or sets the DateTimeStyles used to map the value to a DateTime.
    /// </summary>
    public DateTimeStyles Style { get; set; }

    /// <inheritdoc/>
    public CellMapperResult Map(ReadCellResult readResult)
    {
        // Excel stores dates as numbers (the number of days since 1899-12-30).
        // ExcelDataReader automatically converts these cells to DateTime.
        if (readResult.GetValue() is DateTime dateTimeValue)
        {
            return CellMapperResult.Success(dateTimeValue);
        }

        // A number in a cell without a date format is Excel's serial date (days since 1899-12-30).
        if (ExcelDateUtilities.TryGetDateTime(readResult.GetValue(), out var serialDate))
        {
            return CellMapperResult.Success(serialDate);
        }
        
        var stringValue = readResult.GetString();
        var formats = Formats;
        var provider = Provider;
        var style = Style;
        try
        {
            if (DateTime.TryParseExact(stringValue, formats, provider, style, out var result) ||
                (!_formatsSet && FormatUtilities.HasDigit(stringValue) && DateTime.TryParseExact(stringValue, FormatUtilities.GetDefaultDateTimeFormats(provider), provider, style, out result)))
            {
                return CellMapperResult.Success(result);
            }
        }
        catch (Exception exception)
        {
            return CellMapperResult.Invalid(exception);
        }

        // Only create the exception (by parsing again) if it is needed.
        return InvalidResult(stringValue, formats, provider, style);
    }

    CellConversion ITypedCellMapper<DateTime>.Convert(ReadCellResult cell, out DateTime value)
    {
        var cellValue = cell.GetValue();
        if (cellValue is DateTime dateTimeValue)
        {
            value = dateTimeValue;
            return CellConversion.Success;
        }

        if (ExcelDateUtilities.TryGetDateTime(cellValue, out value))
        {
            return CellConversion.Success;
        }

        var stringValue = cell.GetString();
        var provider = Provider;
        var style = Style;
        try
        {
            return DateTime.TryParseExact(stringValue, Formats, provider, style, out value) ||
                (!_formatsSet && FormatUtilities.HasDigit(stringValue) && DateTime.TryParseExact(stringValue, FormatUtilities.GetDefaultDateTimeFormats(provider), provider, style, out value))
                ? CellConversion.Success
                : CellConversion.Invalid;
        }
        catch
        {
            value = default;
            return CellConversion.Invalid;
        }
    }

    CellConversion ITypedCellMapper<DateTime?>.Convert(ReadCellResult cell, out DateTime? value)
    {
        var conversion = ((ITypedCellMapper<DateTime>)this).Convert(cell, out var result);
        value = conversion == CellConversion.Invalid ? null : result;
        return conversion;
    }

    // A separate method so that the closure capturing these values is only allocated for invalid
    // values. Creating the lambda in Map would allocate the closure for every cell.
    private static CellMapperResult InvalidResult(string? stringValue, string[] formats, IFormatProvider? provider, DateTimeStyles style)
        => MapperUtilities.InvalidFromParse(() => DateTime.ParseExact(stringValue!, formats, provider, style));
}

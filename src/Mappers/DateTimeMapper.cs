using System.Globalization;

namespace ExcelMapper.Mappers;

/// <summary>
/// A mapper that tries to map the value of a cell to a DateTime.
/// </summary>
internal sealed class DateTimeMapper : IFormatsCellMapper, ICellMapper<DateTime>, ICellMapper<DateTime?>
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

    CellConversion ICellMapper<DateTime>.Convert(ReadCellResult cell, out DateTime value)
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
        var provider = Provider ?? cell.Culture;
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

    CellConversion ICellMapper<DateTime?>.Convert(ReadCellResult cell, out DateTime? value)
    {
        var conversion = ((ICellMapper<DateTime>)this).Convert(cell, out var result);
        value = conversion == CellConversion.Invalid ? null : result;
        return conversion;
    }

    public Exception CreateException(ReadCellResult cell)
    {
        var stringValue = cell.GetString();
        var formats = Formats;
        var provider = Provider ?? cell.Culture;
        var style = Style;
        try
        {
            // Parsing throws for some styles, such as invalid combinations.
            _ = DateTime.TryParseExact(stringValue, formats, provider, style, out _) ||
                (!_formatsSet && FormatUtilities.HasDigit(stringValue) && DateTime.TryParseExact(stringValue, FormatUtilities.GetDefaultDateTimeFormats(provider), provider, style, out _));
        }
        catch (Exception exception)
        {
            return exception;
        }

        return MapperUtilities.InvalidDate(stringValue, style, "DateTime", formats, () => DateTime.ParseExact(stringValue!, formats, provider, style));
    }
}

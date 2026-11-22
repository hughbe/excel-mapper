#if NET6_0_OR_GREATER
using System.Globalization;

namespace ExcelMapper.Mappers;

/// <summary>
/// A mapper that tries to map the value of a cell to a TimeOnly.
/// </summary>
internal sealed class TimeOnlyMapper : IFormatsCellMapper, ICellMapper<TimeOnly>, ICellMapper<TimeOnly?>
{
    private string[] _formats = ["t"];

    /// <summary>
    /// Gets or sets the date formats used to map the value to a TimeOnly.
    /// This defaults to "t" - the default TimeOnly format.
    /// While the formats haven't been set, values that don't match also try the culture's long time ("T") and 24-hour times such as "14:30" or "14:30:15".
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
    /// Gets or sets the IFormatProvider used to map the value to a TimeOnly.
    /// </summary>
    public IFormatProvider? Provider { get; set; }

    /// <summary>
    /// Gets or sets the DateTimeStyles used to map the value to a TimeOnly.
    /// </summary>
    public DateTimeStyles Style { get; set; }

    CellConversion ICellMapper<TimeOnly>.Convert(ReadCellResult cell, out TimeOnly value)
    {
        var cellValue = cell.GetValue();
        if (cellValue is TimeSpan timeSpanValue)
        {
            try
            {
                value = TimeOnly.FromTimeSpan(timeSpanValue);
                return CellConversion.Success;
            }
            catch
            {
                value = default;
                return CellConversion.Invalid;
            }
        }

        if (cellValue is DateTime dateTimeValue)
        {
            value = TimeOnly.FromDateTime(dateTimeValue);
            return CellConversion.Success;
        }

        if (cellValue is double fractionOfDay && fractionOfDay >= 0 && fractionOfDay < 1)
        {
            value = new TimeOnly((long)Math.Round(fractionOfDay * TimeSpan.TicksPerDay));
            return CellConversion.Success;
        }

        var stringValue = cell.GetString();
        var provider = Provider ?? cell.Culture;
        var style = Style;
        try
        {
            return TimeOnly.TryParseExact(stringValue, Formats, provider, style, out value) ||
                (!_formatsSet && FormatUtilities.HasDigit(stringValue) && TimeOnly.TryParseExact(stringValue, FormatUtilities.GetDefaultTimeOnlyFormats(provider), provider, style, out value))
                ? CellConversion.Success
                : CellConversion.Invalid;
        }
        catch
        {
            value = default;
            return CellConversion.Invalid;
        }
    }

    CellConversion ICellMapper<TimeOnly?>.Convert(ReadCellResult cell, out TimeOnly? value)
    {
        var conversion = ((ICellMapper<TimeOnly>)this).Convert(cell, out var result);
        value = conversion == CellConversion.Invalid ? null : result;
        return conversion;
    }

    public Exception CreateException(ReadCellResult cell)
    {
        // A duration that isn't a time of day.
        if (cell.GetValue() is TimeSpan timeSpanValue)
        {
            return MapperUtilities.InvalidFromParse(() => TimeOnly.FromTimeSpan(timeSpanValue));
        }

        var stringValue = cell.GetString();
        var formats = Formats;
        var provider = Provider ?? cell.Culture;
        var style = Style;
        try
        {
            // Parsing throws for some styles, such as invalid combinations.
            _ = TimeOnly.TryParseExact(stringValue, formats, provider, style, out _) ||
                (!_formatsSet && FormatUtilities.HasDigit(stringValue) && TimeOnly.TryParseExact(stringValue, FormatUtilities.GetDefaultTimeOnlyFormats(provider), provider, style, out _));
        }
        catch (Exception exception)
        {
            return exception;
        }

        return MapperUtilities.InvalidDate(stringValue, style, "TimeOnly", formats, () => TimeOnly.ParseExact(stringValue!, formats, provider, style));
    }
}
#endif

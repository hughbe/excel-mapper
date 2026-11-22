using System.Globalization;

namespace ExcelMapper.Mappers;

/// <summary>
/// A mapper that tries to map the value of a cell to a DateTimeOffset.
/// </summary>
internal sealed class DateTimeOffsetMapper : IFormatsCellMapper, ICellMapper<DateTimeOffset>, ICellMapper<DateTimeOffset?>
{
    private string[] _formats = ["G"];

    /// <summary>
    /// Gets or sets the date formats used to map the value to a DateTimeOffset.
    /// This defaults to "G" - the default Excel format.
    /// While the formats haven't been set, values that don't match also try the culture's short date ("d") and short date and time ("g"), and ISO 8601, such as "2024-03-04" or "2024-03-04T10:30:00+01:00".
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
    /// Gets or sets the IFormatProvider used to map the value to a DateTimeOffset.
    /// </summary>
    public IFormatProvider? Provider { get; set; }

    /// <summary>
    /// Gets or sets the DateTimeStyles used to map the value to a DateTimeOffset.
    /// </summary>
    public DateTimeStyles Style { get; set; }

    private TimeSpan? _offset;

    /// <summary>
    /// Gets or sets the offset from UTC given to cells that Excel stores as dates. Excel dates have no
    /// time zone, so when this is null (the default) the machine's local offset is used, which makes
    /// results depend on where the code runs. Set to <see cref="TimeSpan.Zero"/> for UTC. Text with
    /// an explicit offset (for example a "zzz" format) keeps that offset.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The offset is not a whole number of minutes
    /// between -14 and +14 hours.</exception>
    public TimeSpan? Offset
    {
        get => _offset;
        set
        {
            if (value is { } offset && (offset.Ticks % TimeSpan.TicksPerMinute != 0 || offset < TimeSpan.FromHours(-14) || offset > TimeSpan.FromHours(14)))
            {
                throw new ArgumentOutOfRangeException(nameof(value), offset, "The offset must be a whole number of minutes between -14 and +14 hours.");
            }

            _offset = value;
        }
    }

    CellConversion ICellMapper<DateTimeOffset>.Convert(ReadCellResult cell, out DateTimeOffset value)
    {
        var cellValue = cell.GetValue();
        if (cellValue is DateTime dateTimeValue || ExcelDateUtilities.TryGetDateTime(cellValue, out dateTimeValue))
        {
            try
            {
                value = Offset is { } offset
                    ? new DateTimeOffset(DateTime.SpecifyKind(dateTimeValue, DateTimeKind.Unspecified), offset)
                    : new DateTimeOffset(dateTimeValue);
                return CellConversion.Success;
            }
            catch (ArgumentOutOfRangeException)
            {
                value = default;
                return CellConversion.Invalid;
            }
        }

        var stringValue = cell.GetString();
        var provider = Provider;
        var style = Style;
        try
        {
            return DateTimeOffset.TryParseExact(stringValue, Formats, provider, style, out value) ||
                (!_formatsSet && FormatUtilities.HasDigit(stringValue) && DateTimeOffset.TryParseExact(stringValue, FormatUtilities.GetDefaultDateTimeFormats(provider), provider, style, out value))
                ? CellConversion.Success
                : CellConversion.Invalid;
        }
        catch
        {
            value = default;
            return CellConversion.Invalid;
        }
    }

    CellConversion ICellMapper<DateTimeOffset?>.Convert(ReadCellResult cell, out DateTimeOffset? value)
    {
        var conversion = ((ICellMapper<DateTimeOffset>)this).Convert(cell, out var result);
        value = conversion == CellConversion.Invalid ? null : result;
        return conversion;
    }

    public Exception CreateException(ReadCellResult cell)
    {
        // A date that can't have the offset, such as the earliest dates with a positive offset.
        var cellValue = cell.GetValue();
        if (cellValue is DateTime dateTimeValue || ExcelDateUtilities.TryGetDateTime(cellValue, out dateTimeValue))
        {
            var offset = Offset;
            return MapperUtilities.InvalidFromParse(() => _ = offset is { } value
                ? new DateTimeOffset(DateTime.SpecifyKind(dateTimeValue, DateTimeKind.Unspecified), value)
                : new DateTimeOffset(dateTimeValue));
        }

        var stringValue = cell.GetString();
        var formats = Formats;
        var provider = Provider;
        var style = Style;
        try
        {
            // Parsing throws for some styles, such as invalid combinations.
            _ = DateTimeOffset.TryParseExact(stringValue, formats, provider, style, out _) ||
                (!_formatsSet && FormatUtilities.HasDigit(stringValue) && DateTimeOffset.TryParseExact(stringValue, FormatUtilities.GetDefaultDateTimeFormats(provider), provider, style, out _));
        }
        catch (Exception exception)
        {
            return exception;
        }

        return MapperUtilities.InvalidDate(stringValue, style, "DateTimeOffset", formats, () => DateTimeOffset.ParseExact(stringValue!, formats, provider, style));
    }
}

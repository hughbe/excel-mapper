using System.Globalization;

namespace ExcelMapper.Mappers;

/// <summary>
/// A mapper that tries to map the value of a cell to a DateTimeOffset.
/// </summary>
public class DateTimeOffsetMapper : IFormatsCellMapper
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

    private CellMapperResult CreateResult(DateTime dateTime)
    {
        try
        {
            return CellMapperResult.Success(Offset is { } offset
                ? new DateTimeOffset(DateTime.SpecifyKind(dateTime, DateTimeKind.Unspecified), offset)
                : new DateTimeOffset(dateTime));
        }
        catch (ArgumentOutOfRangeException exception)
        {
            // For example, the earliest dates with a positive offset.
            return CellMapperResult.Invalid(exception);
        }
    }

    /// <inheritdoc/>
    public CellMapperResult Map(ReadCellResult readResult)
    {
        // Excel stores dates as numbers (the number of days since 1899-12-30).
        // ExcelDataReader automatically converts these cells to DateTimeOffset.
        if (readResult.GetValue() is DateTime dateTimeOffsetValue)
        {
            return CreateResult(dateTimeOffsetValue);
        }

        // A number in a cell without a date format is Excel's serial date (days since 1899-12-30).
        if (ExcelDateUtilities.TryGetDateTime(readResult.GetValue(), out var serialDate))
        {
            return CreateResult(serialDate);
        }
        
        var stringValue = readResult.GetString();
        var formats = Formats;
        var provider = Provider;
        var style = Style;
        try
        {
            if (DateTimeOffset.TryParseExact(stringValue, formats, provider, style, out var result) ||
                (!_formatsSet && FormatUtilities.HasDigit(stringValue) && DateTimeOffset.TryParseExact(stringValue, FormatUtilities.GetDefaultDateTimeFormats(provider), provider, style, out result)))
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

    // A separate method so that the closure capturing these values is only allocated for invalid
    // values. Creating the lambda in Map would allocate the closure for every cell.
    private static CellMapperResult InvalidResult(string? stringValue, string[] formats, IFormatProvider? provider, DateTimeStyles style)
        => MapperUtilities.InvalidFromParse(() => DateTimeOffset.ParseExact(stringValue!, formats, provider, style));
}

using System.Globalization;

namespace ExcelMapper.Mappers;

/// <summary>
/// A mapper that tries to map the value of a cell to a TimeSpan.
/// </summary>
public class TimeSpanMapper : ICellMapper, IFormatsCellMapper
{
    private string[] _formats = ["c"];

    /// <summary>
    /// Gets or sets the date formats used to map the value to a TimeSpan.
    /// This defaults to "c" - the default TimeSpan format.
    /// </summary>
    public string[] Formats
    {
        get => _formats;
        set
        {
            FormatUtilities.ValidateFormats(value);
            _formats = value;
        }
    }

    /// <summary>
    /// Gets or sets the IFormatProvider used to map the value to a TimeSpan.
    /// </summary>
    public IFormatProvider? Provider { get; set; }

    /// <summary>
    /// Gets or sets the TimeSpanStyles used to map the value to a TimeSpan.
    /// </summary>
    public TimeSpanStyles Style { get; set; }

    /// <inheritdoc/>
    public CellMapperResult Map(ReadCellResult readResult)
    {
        // Excel stores durations as number of days.
        // ExcelDataReader automatically converts these cells to TimeSpan.
        if (readResult.GetValue() is TimeSpan timeSpanValue)
        {
            return CellMapperResult.Success(timeSpanValue);
        }

        // Cells formatted as a time of day (rather than a duration) are read as a DateTime on Excel's
        // first day.
        if (readResult.GetValue() is DateTime dateTimeValue && ExcelDateUtilities.IsTimeOfDayOnly(dateTimeValue))
        {
            return CellMapperResult.Success(dateTimeValue.TimeOfDay);
        }

        // A number in a cell without a time format is a number of days, as Excel stores durations.
        if (readResult.GetValue() is double days && !double.IsNaN(days) && Math.Abs(days) <= TimeSpan.MaxValue.TotalDays)
        {
            return CellMapperResult.Success(TimeSpan.FromTicks((long)Math.Round(days * TimeSpan.TicksPerDay)));
        }
        
        var stringValue = readResult.GetString();
        var formats = Formats;
        var provider = Provider;
        var style = Style;
        try
        {
            if (TimeSpan.TryParseExact(stringValue, formats, provider, style, out var result))
            {
                return CellMapperResult.Success(result);
            }
        }
        catch (Exception exception)
        {
            return CellMapperResult.Invalid(exception);
        }

        // Only create the exception (by parsing again) if it is needed.
        return MapperUtilities.InvalidFromParse(() => TimeSpan.ParseExact(stringValue!, formats, provider, style));
    }
}

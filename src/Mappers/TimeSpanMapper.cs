using System.Globalization;

namespace ExcelMapper.Mappers;

/// <summary>
/// A mapper that tries to map the value of a cell to a TimeSpan.
/// </summary>
internal sealed class TimeSpanMapper : IFormatsCellMapper, ICellMapper<TimeSpan>, ICellMapper<TimeSpan?>
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

    CellConversion ICellMapper<TimeSpan>.Convert(ReadCellResult cell, out TimeSpan value)
    {
        var cellValue = cell.GetValue();
        if (cellValue is TimeSpan timeSpanValue)
        {
            value = timeSpanValue;
            return CellConversion.Success;
        }

        if (cellValue is DateTime dateTimeValue && ExcelDateUtilities.IsTimeOfDayOnly(dateTimeValue))
        {
            value = dateTimeValue.TimeOfDay;
            return CellConversion.Success;
        }

        if (cellValue is double days && !double.IsNaN(days) && Math.Abs(days) <= TimeSpan.MaxValue.TotalDays)
        {
            value = TimeSpan.FromTicks((long)Math.Round(days * TimeSpan.TicksPerDay));
            return CellConversion.Success;
        }

        try
        {
            return TimeSpan.TryParseExact(cell.GetString(), Formats, Provider ?? cell.Culture, Style, out value)
                ? CellConversion.Success
                : CellConversion.Invalid;
        }
        catch
        {
            value = default;
            return CellConversion.Invalid;
        }
    }

    CellConversion ICellMapper<TimeSpan?>.Convert(ReadCellResult cell, out TimeSpan? value)
    {
        var conversion = ((ICellMapper<TimeSpan>)this).Convert(cell, out var result);
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
            _ = TimeSpan.TryParseExact(stringValue, formats, provider, style, out _);
        }
        catch (Exception exception)
        {
            return exception;
        }

        return MapperUtilities.InvalidFromParse(() => TimeSpan.ParseExact(stringValue!, formats, provider, style));
    }
}

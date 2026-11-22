using System.Globalization;

namespace ExcelMapper.Mappers;

/// <summary>
/// A mapper that tries to map the value of a cell to a TimeSpan.
/// </summary>
internal sealed class TimeSpanMapper : ICellMapper<TimeSpan>, ICellMapper<TimeSpan?>
{
    /// <summary>
    /// The mapper, shared by every member: the formats and format provider come from the cell's
    /// <see cref="CellContext.Options"/>.
    /// </summary>
    public static TimeSpanMapper Instance { get; } = new();

    private TimeSpanMapper()
    {
    }

    // The format tried first when the member has no formats of its own.
    private static readonly string[] s_defaultFormats = ["c"];

    private const TimeSpanStyles Style = TimeSpanStyles.None;

    CellConversion ICellMapper<TimeSpan>.Convert(ReadCellResult cell, in CellContext context, out TimeSpan value, ref object? error)
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
            return TimeSpan.TryParseExact(cell.GetString(), (context.Options.Formats ?? s_defaultFormats), context.Options.Provider ?? cell.Culture, Style, out value)
                ? CellConversion.Success
                : CellConversion.Invalid;
        }
        catch
        {
            value = default;
            return CellConversion.Invalid;
        }
    }

    CellConversion ICellMapper<TimeSpan?>.Convert(ReadCellResult cell, in CellContext context, out TimeSpan? value, ref object? error)
    {
        var conversion = ((ICellMapper<TimeSpan>)this).Convert(cell, in context, out var result, ref error);
        value = conversion == CellConversion.Invalid ? null : result;
        return conversion;
    }

    public Exception CreateException(ReadCellResult cell, in CellContext context, object? error)
    {
        var stringValue = cell.GetString();
        var formats = context.Options.Formats ?? s_defaultFormats;
        var provider = context.Options.Provider ?? cell.Culture;
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

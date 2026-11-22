#if NET6_0_OR_GREATER
using System.Globalization;

namespace ExcelMapper.Mappers;

/// <summary>
/// A mapper that tries to map the value of a cell to a TimeOnly.
/// </summary>
internal sealed class TimeOnlyMapper : ICellMapper<TimeOnly>, ICellMapper<TimeOnly?>
{
    /// <summary>
    /// The mapper, shared by every member: the formats and format provider come from the cell's
    /// <see cref="CellContext.Options"/>.
    /// </summary>
    public static TimeOnlyMapper Instance { get; } = new();

    private TimeOnlyMapper()
    {
    }

    // The format tried first when the member has no formats of its own.
    private static readonly string[] s_defaultFormats = ["t"];

    private const DateTimeStyles Style = DateTimeStyles.None;

    CellConversion ICellMapper<TimeOnly>.Convert(ReadCellResult cell, in CellContext context, out TimeOnly value, ref object? error)
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
        var provider = context.Options.Provider ?? cell.Culture;
        var style = Style;
        try
        {
            return TimeOnly.TryParseExact(stringValue, (context.Options.Formats ?? s_defaultFormats), provider, style, out value) ||
                (context.Options.Formats is null && FormatUtilities.HasDigit(stringValue) && TimeOnly.TryParseExact(stringValue, FormatUtilities.GetDefaultTimeOnlyFormats(provider), provider, style, out value))
                ? CellConversion.Success
                : CellConversion.Invalid;
        }
        catch
        {
            value = default;
            return CellConversion.Invalid;
        }
    }

    CellConversion ICellMapper<TimeOnly?>.Convert(ReadCellResult cell, in CellContext context, out TimeOnly? value, ref object? error)
    {
        var conversion = ((ICellMapper<TimeOnly>)this).Convert(cell, in context, out var result, ref error);
        value = conversion == CellConversion.Invalid ? null : result;
        return conversion;
    }

    public Exception CreateException(ReadCellResult cell, in CellContext context, object? error)
    {
        // A duration that isn't a time of day.
        if (cell.GetValue() is TimeSpan timeSpanValue)
        {
            return MapperUtilities.InvalidFromParse(() => TimeOnly.FromTimeSpan(timeSpanValue));
        }

        var stringValue = cell.GetString();
        var formats = context.Options.Formats ?? s_defaultFormats;
        var provider = context.Options.Provider ?? cell.Culture;
        var style = Style;
        try
        {
            // Parsing throws for some styles, such as invalid combinations.
            _ = TimeOnly.TryParseExact(stringValue, formats, provider, style, out _) ||
                (context.Options.Formats is null && FormatUtilities.HasDigit(stringValue) && TimeOnly.TryParseExact(stringValue, FormatUtilities.GetDefaultTimeOnlyFormats(provider), provider, style, out _));
        }
        catch (Exception exception)
        {
            return exception;
        }

        return MapperUtilities.InvalidDate(stringValue, style, "TimeOnly", formats, () => TimeOnly.ParseExact(stringValue!, formats, provider, style));
    }
}
#endif

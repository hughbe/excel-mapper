using System.Globalization;

namespace ExcelMapper.Mappers;

/// <summary>
/// A mapper that tries to map the value of a cell to a DateTime.
/// </summary>
internal sealed class DateTimeMapper : ICellMapper<DateTime>, ICellMapper<DateTime?>
{
    /// <summary>
    /// The mapper, shared by every member: the formats and format provider come from the cell's
    /// <see cref="CellContext.Options"/>.
    /// </summary>
    public static DateTimeMapper Instance { get; } = new();

    private DateTimeMapper()
    {
    }

    // The format tried first when the member has no formats of its own.
    private static readonly string[] s_defaultFormats = ["G"];

    // Text with an offset ("Z", "+01:00") is converted to UTC, rather than to the machine's local
    // time, so a file gives the same values on every machine. Text without one is left as it is.
    private const DateTimeStyles Style = DateTimeStyles.AdjustToUniversal;

    CellConversion ICellMapper<DateTime>.Convert(ReadCellResult cell, in CellContext context, out DateTime value, ref object? error)
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
        var provider = context.Options.Provider ?? cell.Culture;
        var style = Style;
        try
        {
            // ISO 8601 text, the most common in CSV files, is parsed with the one default format
            // it can match before the culture's patterns, which is several times faster.
            if (context.Options.Formats is null && FormatUtilities.GetIsoDateTimeFormat(stringValue) is { } isoFormat &&
                DateTime.TryParseExact(stringValue, isoFormat, provider, style, out value))
            {
                return CellConversion.Success;
            }

            return DateTime.TryParseExact(stringValue, (context.Options.Formats ?? s_defaultFormats), provider, style, out value) ||
                (context.Options.Formats is null && FormatUtilities.HasDigit(stringValue) && DateTime.TryParseExact(stringValue, FormatUtilities.GetDefaultDateTimeFormats(provider), provider, style, out value))
                ? CellConversion.Success
                : CellConversion.Invalid;
        }
        catch
        {
            value = default;
            return CellConversion.Invalid;
        }
    }

    CellConversion ICellMapper<DateTime?>.Convert(ReadCellResult cell, in CellContext context, out DateTime? value, ref object? error)
    {
        var conversion = ((ICellMapper<DateTime>)this).Convert(cell, in context, out var result, ref error);
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
            _ = DateTime.TryParseExact(stringValue, formats, provider, style, out _) ||
                (context.Options.Formats is null && FormatUtilities.HasDigit(stringValue) && DateTime.TryParseExact(stringValue, FormatUtilities.GetDefaultDateTimeFormats(provider), provider, style, out _));
        }
        catch (Exception exception)
        {
            return exception;
        }

        return MapperUtilities.InvalidDate(stringValue, style, "DateTime", formats, () => DateTime.ParseExact(stringValue!, formats, provider, style));
    }
}

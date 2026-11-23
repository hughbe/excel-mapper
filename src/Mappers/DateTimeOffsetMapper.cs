using System.Globalization;

namespace ExcelMapper.Mappers;

/// <summary>
/// A mapper that tries to map the value of a cell to a DateTimeOffset.
/// </summary>
internal sealed class DateTimeOffsetMapper : ICellMapper<DateTimeOffset>, ICellMapper<DateTimeOffset?>
{
    /// <summary>
    /// The mapper, shared by every member: the formats and format provider come from the cell's
    /// <see cref="CellContext.Options"/>.
    /// </summary>
    public static DateTimeOffsetMapper Instance { get; } = new();

    private DateTimeOffsetMapper()
    {
    }

    // The format tried first when the member has no formats of its own.
    private static readonly string[] s_defaultFormats = ["G"];

    // A date without an offset gets the member's offset (WithOffset), or UTC, rather than the
    // machine's local offset, so a file gives the same values on every machine.
    private const DateTimeStyles Style = DateTimeStyles.AssumeUniversal;

    CellConversion ICellMapper<DateTimeOffset>.Convert(ReadCellResult cell, in CellContext context, out DateTimeOffset value, ref object? error)
    {
        var cellValue = cell.GetValue();
        if (cellValue is DateTime dateTimeValue || ExcelDateUtilities.TryGetDateTime(cellValue, out dateTimeValue))
        {
            try
            {
                value = new DateTimeOffset(DateTime.SpecifyKind(dateTimeValue, DateTimeKind.Unspecified), context.Options.Offset ?? TimeSpan.Zero);
                return CellConversion.Success;
            }
            catch (ArgumentOutOfRangeException)
            {
                value = default;
                return CellConversion.Invalid;
            }
        }

        var stringValue = cell.GetString();
        var provider = context.Options.Provider ?? cell.Culture;
        var style = Style;
        try
        {
            // Text without an offset of its own gets the member's offset. Parsed as a DateTime
            // first, whose kind says whether the text had one: only for members with an offset.
            if (context.Options.Offset is { } memberOffset &&
                TryParseDateTime(stringValue, context, provider, out var dateTime) &&
                dateTime.Kind == DateTimeKind.Unspecified)
            {
                value = new DateTimeOffset(dateTime, memberOffset);
                return CellConversion.Success;
            }

            // ISO 8601 text, the most common in CSV files, is parsed with the one default format
            // it can match before the culture's patterns, which is several times faster.
            if (context.Options.Formats is null && FormatUtilities.GetIsoDateTimeFormat(stringValue) is { } isoFormat &&
                DateTimeOffset.TryParseExact(stringValue, isoFormat, provider, style, out value))
            {
                return CellConversion.Success;
            }

            return DateTimeOffset.TryParseExact(stringValue, (context.Options.Formats ?? s_defaultFormats), provider, style, out value) ||
                (context.Options.Formats is null && FormatUtilities.HasDigit(stringValue) && DateTimeOffset.TryParseExact(stringValue, FormatUtilities.GetDefaultDateTimeFormats(provider), provider, style, out value))
                ? CellConversion.Success
                : CellConversion.Invalid;
        }
        catch
        {
            value = default;
            return CellConversion.Invalid;
        }
    }

    // Parses the text as a DateTimeOffset member's formats would, as a DateTime whose kind is
    // Unspecified if the text has no offset.
    private static bool TryParseDateTime(string? text, in CellContext context, IFormatProvider provider, out DateTime value)
        => (context.Options.Formats is null && FormatUtilities.GetIsoDateTimeFormat(text) is { } isoFormat &&
                DateTime.TryParseExact(text, isoFormat, provider, DateTimeStyles.None, out value)) ||
            DateTime.TryParseExact(text, context.Options.Formats ?? s_defaultFormats, provider, DateTimeStyles.None, out value) ||
            (context.Options.Formats is null && FormatUtilities.HasDigit(text) && DateTime.TryParseExact(text, FormatUtilities.GetDefaultDateTimeFormats(provider), provider, DateTimeStyles.None, out value));

    CellConversion ICellMapper<DateTimeOffset?>.Convert(ReadCellResult cell, in CellContext context, out DateTimeOffset? value, ref object? error)
    {
        var conversion = ((ICellMapper<DateTimeOffset>)this).Convert(cell, in context, out var result, ref error);
        value = conversion == CellConversion.Invalid ? null : result;
        return conversion;
    }

    public Exception CreateException(ReadCellResult cell, in CellContext context, object? error)
    {
        // A date that can't have the offset, such as the earliest dates with a positive offset.
        var cellValue = cell.GetValue();
        if (cellValue is DateTime dateTimeValue || ExcelDateUtilities.TryGetDateTime(cellValue, out dateTimeValue))
        {
            var offset = context.Options.Offset ?? TimeSpan.Zero;
            return MapperUtilities.InvalidFromParse(() => _ = new DateTimeOffset(DateTime.SpecifyKind(dateTimeValue, DateTimeKind.Unspecified), offset));
        }

        var stringValue = cell.GetString();
        var formats = context.Options.Formats ?? s_defaultFormats;
        var provider = context.Options.Provider ?? cell.Culture;
        var style = Style;
        try
        {
            // Parsing throws for some styles, such as invalid combinations.
            _ = DateTimeOffset.TryParseExact(stringValue, formats, provider, style, out _) ||
                (context.Options.Formats is null && FormatUtilities.HasDigit(stringValue) && DateTimeOffset.TryParseExact(stringValue, FormatUtilities.GetDefaultDateTimeFormats(provider), provider, style, out _));
        }
        catch (Exception exception)
        {
            return exception;
        }

        return MapperUtilities.InvalidDate(stringValue, style, "DateTimeOffset", formats, () => DateTimeOffset.ParseExact(stringValue!, formats, provider, style));
    }
}

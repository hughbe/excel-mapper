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

    private const DateTimeStyles Style = DateTimeStyles.None;

    CellConversion ICellMapper<DateTimeOffset>.Convert(ReadCellResult cell, in CellContext context, out DateTimeOffset value, ref object? error)
    {
        var cellValue = cell.GetValue();
        if (cellValue is DateTime dateTimeValue || ExcelDateUtilities.TryGetDateTime(cellValue, out dateTimeValue))
        {
            try
            {
                value = context.Options.Offset is { } offset
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
        var provider = context.Options.Provider ?? cell.Culture;
        var style = Style;
        try
        {
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
            var offset = context.Options.Offset;
            return MapperUtilities.InvalidFromParse(() => _ = offset is { } value
                ? new DateTimeOffset(DateTime.SpecifyKind(dateTimeValue, DateTimeKind.Unspecified), value)
                : new DateTimeOffset(dateTimeValue));
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

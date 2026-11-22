#if NET6_0_OR_GREATER
using System.Globalization;

namespace ExcelMapper.Mappers;

/// <summary>
/// A mapper that tries to map the value of a cell to a DateOnly.
/// </summary>
internal sealed class DateOnlyMapper : ICellMapper<DateOnly>, ICellMapper<DateOnly?>
{
    /// <summary>
    /// The mapper, shared by every member: the formats and format provider come from the cell's
    /// <see cref="CellContext.Options"/>.
    /// </summary>
    public static DateOnlyMapper Instance { get; } = new();

    private DateOnlyMapper()
    {
    }

    // The format tried first when the member has no formats of its own.
    private static readonly string[] s_defaultFormats = ["d"];

    private const DateTimeStyles Style = DateTimeStyles.None;

    CellConversion ICellMapper<DateOnly>.Convert(ReadCellResult cell, in CellContext context, out DateOnly value, ref object? error)
    {
        var cellValue = cell.GetValue();
        if (cellValue is DateTime dateTimeValue || ExcelDateUtilities.TryGetDateTime(cellValue, out dateTimeValue))
        {
            value = DateOnly.FromDateTime(dateTimeValue);
            return CellConversion.Success;
        }

        var stringValue = cell.GetString();
        var provider = context.Options.Provider ?? cell.Culture;
        var style = Style;
        try
        {
            return DateOnly.TryParseExact(stringValue, (context.Options.Formats ?? s_defaultFormats), provider, style, out value) ||
                (context.Options.Formats is null && FormatUtilities.HasDigit(stringValue) && DateOnly.TryParseExact(stringValue, FormatUtilities.DefaultDateOnlyFormats, provider, style, out value))
                ? CellConversion.Success
                : CellConversion.Invalid;
        }
        catch
        {
            value = default;
            return CellConversion.Invalid;
        }
    }

    CellConversion ICellMapper<DateOnly?>.Convert(ReadCellResult cell, in CellContext context, out DateOnly? value, ref object? error)
    {
        var conversion = ((ICellMapper<DateOnly>)this).Convert(cell, in context, out var result, ref error);
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
            _ = DateOnly.TryParseExact(stringValue, formats, provider, style, out _) ||
                (context.Options.Formats is null && FormatUtilities.HasDigit(stringValue) && DateOnly.TryParseExact(stringValue, FormatUtilities.DefaultDateOnlyFormats, provider, style, out _));
        }
        catch (Exception exception)
        {
            return exception;
        }

        return MapperUtilities.InvalidDate(stringValue, style, "DateOnly", formats, () => DateOnly.ParseExact(stringValue!, formats, provider, style));
    }
}
#endif

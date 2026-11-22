#if NET6_0_OR_GREATER
using System.Globalization;

namespace ExcelMapper.Mappers;

/// <summary>
/// A mapper that tries to map the value of a cell to a DateOnly.
/// </summary>
internal sealed class DateOnlyMapper : IFormatsCellMapper, ICellMapper<DateOnly>, ICellMapper<DateOnly?>
{
    private string[] _formats = ["d"];

    /// <summary>
    /// Gets or sets the date formats used to map the value to a DateOnly.
    /// This defaults to "d" - the default DateOnly format.
    /// While the formats haven't been set, values that don't match also try ISO 8601, such as "2024-03-04".
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
    /// Gets or sets the IFormatProvider used to map the value to a DateOnly.
    /// </summary>
    public IFormatProvider? Provider { get; set; }

    /// <summary>
    /// Gets or sets the DateTimeStyles used to map the value to a DateOnly.
    /// </summary>
    public DateTimeStyles Style { get; set; }

    CellConversion ICellMapper<DateOnly>.Convert(ReadCellResult cell, in CellContext context, out DateOnly value, ref object? error)
    {
        var cellValue = cell.GetValue();
        if (cellValue is DateTime dateTimeValue || ExcelDateUtilities.TryGetDateTime(cellValue, out dateTimeValue))
        {
            value = DateOnly.FromDateTime(dateTimeValue);
            return CellConversion.Success;
        }

        var stringValue = cell.GetString();
        var provider = Provider ?? cell.Culture;
        var style = Style;
        try
        {
            return DateOnly.TryParseExact(stringValue, Formats, provider, style, out value) ||
                (!_formatsSet && FormatUtilities.HasDigit(stringValue) && DateOnly.TryParseExact(stringValue, FormatUtilities.DefaultDateOnlyFormats, provider, style, out value))
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

    public Exception CreateException(ReadCellResult cell, object? error)
    {
        var stringValue = cell.GetString();
        var formats = Formats;
        var provider = Provider ?? cell.Culture;
        var style = Style;
        try
        {
            // Parsing throws for some styles, such as invalid combinations.
            _ = DateOnly.TryParseExact(stringValue, formats, provider, style, out _) ||
                (!_formatsSet && FormatUtilities.HasDigit(stringValue) && DateOnly.TryParseExact(stringValue, FormatUtilities.DefaultDateOnlyFormats, provider, style, out _));
        }
        catch (Exception exception)
        {
            return exception;
        }

        return MapperUtilities.InvalidDate(stringValue, style, "DateOnly", formats, () => DateOnly.ParseExact(stringValue!, formats, provider, style));
    }
}
#endif

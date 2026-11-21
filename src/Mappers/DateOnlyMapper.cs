#if NET6_0_OR_GREATER
using System.Globalization;

namespace ExcelMapper.Mappers;

/// <summary>
/// A mapper that tries to map the value of a cell to a DateOnly.
/// </summary>
public class DateOnlyMapper : ICellMapper, IFormatsCellMapper
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

    /// <inheritdoc/>
    public CellMapperResult Map(ReadCellResult readResult)
    {
        // Excel stores dates as numbers (the number of days since 1899-12-30).
        // ExcelDataReader automatically converts these cells to DateOnly.
        if (readResult.GetValue() is DateTime dateTimeValue)
        {
            return CellMapperResult.Success(DateOnly.FromDateTime(dateTimeValue));
        }

        // A number in a cell without a date format is Excel's serial date (days since 1899-12-30).
        if (ExcelDateUtilities.TryGetDateTime(readResult.GetValue(), out var serialDate))
        {
            return CellMapperResult.Success(DateOnly.FromDateTime(serialDate));
        }
        
        var stringValue = readResult.GetString();
        var formats = Formats;
        var provider = Provider;
        var style = Style;
        try
        {
            if (DateOnly.TryParseExact(stringValue, formats, provider, style, out var result) ||
                (!_formatsSet && FormatUtilities.HasDigit(stringValue) && DateOnly.TryParseExact(stringValue, FormatUtilities.DefaultDateOnlyFormats, provider, style, out result)))
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
        => MapperUtilities.InvalidFromParse(() => DateOnly.ParseExact(stringValue!, formats, provider, style));
}
#endif

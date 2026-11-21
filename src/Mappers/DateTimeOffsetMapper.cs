using System.Globalization;

namespace ExcelMapper.Mappers;

/// <summary>
/// A mapper that tries to map the value of a cell to a DateTimeOffset.
/// </summary>
public class DateTimeOffsetMapper : IFormatsCellMapper
{
    private string[] _formats = ["G"];

    /// <summary>
    /// Gets or sets the date formats used to map the value to a DateTimeOffset.
    /// This defaults to "G" - the default Excel format.
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
    /// Gets or sets the IFormatProvider used to map the value to a DateTimeOffset.
    /// </summary>
    public IFormatProvider? Provider { get; set; }

    /// <summary>
    /// Gets or sets the DateTimeStyles used to map the value to a DateTimeOffset.
    /// </summary>
    public DateTimeStyles Style { get; set; }

    /// <inheritdoc/>
    public CellMapperResult Map(ReadCellResult readResult)
    {
        // Excel stores dates as numbers (the number of days since 1899-12-30).
        // ExcelDataReader automatically converts these cells to DateTimeOffset.
        if (readResult.GetValue() is DateTime dateTimeOffsetValue)
        {
            return CellMapperResult.Success(new DateTimeOffset(dateTimeOffsetValue));
        }

        // A number in a cell without a date format is Excel's serial date (days since 1899-12-30).
        if (ExcelDateUtilities.TryGetDateTime(readResult.GetValue(), out var serialDate))
        {
            return CellMapperResult.Success(new DateTimeOffset(serialDate));
        }
        
        var stringValue = readResult.GetString();
        var formats = Formats;
        var provider = Provider;
        var style = Style;
        try
        {
            if (DateTimeOffset.TryParseExact(stringValue, formats, provider, style, out var result))
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
        => MapperUtilities.InvalidFromParse(() => DateTimeOffset.ParseExact(stringValue!, formats, provider, style));
}

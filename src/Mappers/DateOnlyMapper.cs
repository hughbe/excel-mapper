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
        
        var stringValue = readResult.GetString();
        try
        {
            var result = DateOnly.ParseExact(stringValue!, Formats, Provider, Style);
            return CellMapperResult.Success(result);
        }
        catch (Exception exception)
        {
            return CellMapperResult.Invalid(exception);
        }
    }
}

#if NET6_0_OR_GREATER
using System.Globalization;

namespace ExcelMapper.Mappers;

/// <summary>
/// A mapper that tries to map the value of a cell to a TimeOnly.
/// </summary>
public class TimeOnlyMapper : ICellMapper, IFormatsCellMapper
{
    private string[] _formats = ["t"];

    /// <summary>
    /// Gets or sets the date formats used to map the value to a TimeOnly.
    /// This defaults to "t" - the default TimeOnly format.
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
    /// Gets or sets the IFormatProvider used to map the value to a TimeOnly.
    /// </summary>
    public IFormatProvider? Provider { get; set; }

    /// <summary>
    /// Gets or sets the DateTimeStyles used to map the value to a TimeOnly.
    /// </summary>
    public DateTimeStyles Style { get; set; }

    /// <inheritdoc/>
    public CellMapperResult Map(ReadCellResult readResult)
    {
        // Excel stores dates as numbers (the number of days since 1899-12-30).
        // ExcelDataReader automatically converts these cells to TimeOnly.
        if (readResult.GetValue() is TimeSpan timeSpanValue)
        {
            try
            {
                return CellMapperResult.Success(TimeOnly.FromTimeSpan(timeSpanValue));
            }
            catch (Exception exception)
            {
                return CellMapperResult.Invalid(exception);
            }
        }
        if (readResult.GetValue() is DateTime dateTimeValue)
        {
            return CellMapperResult.Success(TimeOnly.FromDateTime(dateTimeValue));
        }
        
        var stringValue = readResult.GetString();
        try
        {
            var result = TimeOnly.ParseExact(stringValue!, Formats, Provider, Style);
            return CellMapperResult.Success(result);
        }
        catch (Exception exception)
        {
            return CellMapperResult.Invalid(exception);
        }
    }
}
#endif

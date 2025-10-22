using System.Globalization;

namespace ExcelMapper.Mappers;

/// <summary>
/// A mapper that tries to map the value of a cell to a TimeSpan.
/// </summary>
public class TimeSpanMapper : ICellMapper, IFormatsCellMapper
{
    private string[] _formats = ["c"];

    /// <summary>
    /// Gets or sets the date formats used to map the value to a TimeSpan.
    /// This defaults to "c" - the default TimeSpan format.
    /// </summary>
    public string[] Formats
    {
        get => _formats;
        set
        {
            ArgumentNullException.ThrowIfNull(value);
            if (value.Length == 0)
            {
                throw new ArgumentException("Formats cannot be empty.", nameof(value));
            }

            foreach (var format in value)
            {
                if (string.IsNullOrEmpty(format))
                {
                    throw new ArgumentException("Formats cannot contain null or empty values.", nameof(value));
                }
            }

            _formats = value;
        }
    }

    /// <summary>
    /// Gets or sets the IFormatProvider used to map the value to a TimeSpan.
    /// </summary>
    public IFormatProvider? Provider { get; set; }

    /// <summary>
    /// Gets or sets the TimeSpanStyles used to map the value to a TimeSpan.
    /// </summary>
    public TimeSpanStyles Style { get; set; }

    public CellMapperResult Map(ReadCellResult readResult)
    {
        // Excel stores durations as number of days.
        // ExcelDataReader automatically converts these cells to TimeSpan.
        if (readResult.GetValue() is TimeSpan timeSpanValue)
        {
            return CellMapperResult.Success(timeSpanValue);
        }
        
        var stringValue = readResult.GetString();
        try
        {
            var result = TimeSpan.ParseExact(stringValue!, Formats, Provider, Style);
            return CellMapperResult.Success(result);
        }
        catch (Exception exception)
        {
            return CellMapperResult.Invalid(exception);
        }
    }
}

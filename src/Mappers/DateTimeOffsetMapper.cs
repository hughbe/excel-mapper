using System;
using System.Globalization;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Mappers;

/// <summary>
/// A mapper that tries to map the value of a cell to a DateTimeOffset.
/// </summary>
public class DateTimeOffsetMapper : ICellMapper
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
    /// Gets or sets the IFormatProvider used to map the value to a DateTimeOffset.
    /// </summary>
    public IFormatProvider? Provider { get; set; }

    /// <summary>
    /// Gets or sets the DateTimeStyles used to map the value to a DateTimeOffset.
    /// </summary>
    public DateTimeStyles Style { get; set; }

    public CellMapperResult MapCellValue(ReadCellResult readResult)
    {
        // Excel stores dates as numbers (the number of days since 1899-12-30).
        // ExcelDataReader automatically converts these cells to DateTimeOffset.
        if (readResult.GetValue() is DateTime dateTimeOffsetValue)
        {
            return CellMapperResult.Success(new DateTimeOffset(dateTimeOffsetValue));
        }
        
        var stringValue = readResult.GetString();
        try
        {
            var result = DateTimeOffset.ParseExact(stringValue!, Formats, Provider, Style);
            return CellMapperResult.Success(result);
        }
        catch (Exception exception)
        {
            return CellMapperResult.Invalid(exception);
        }
    }
}

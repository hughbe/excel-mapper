using System;
using System.Globalization;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Mappers;

/// <summary>
/// A mapper that tries to map the value of a cell to a DateTime.
/// </summary>
public class DateTimeMapper : ICellMapper
{
    private string[] _formats = ["G"];

    /// <summary>
    /// Gets or sets the date formats used to map the value to a DateTime.
    /// This defaults to "G" - the default Excel format.
    /// </summary>
    public string[] Formats
    {
        get => _formats;
        set
        {
            if (value == null)
            {
                throw new ArgumentNullException(nameof(value));
            }

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
    /// Gets or sets the IFormatProvider used to map the value to a DateTime.
    /// </summary>
    public IFormatProvider? Provider { get; set; }

    /// <summary>
    /// Gets or sets the DateTimeStyles used to map the value to a DateTime.
    /// </summary>
    public DateTimeStyles Style { get; set; }

    public CellMapperResult MapCellValue(ReadCellResult readResult)
    {
        if (readResult.GetValue() is DateTime dateTimeValue)
        {
            return CellMapperResult.Success(dateTimeValue);
        }
        
        var stringValue = readResult.GetString();
        try
        {
            var result = DateTime.ParseExact(stringValue!, Formats, Provider, Style);
            return CellMapperResult.Success(result);
        }
        catch (Exception exception)
        {
            return CellMapperResult.Invalid(exception);
        }
    }
}

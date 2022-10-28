﻿using System.Globalization;
using System.Reflection;
using ExcelMapper.Abstractions;

/// <summary>
/// A mapper that tries to map the value of a cell to a DateTime.
/// </summary>
public class DateTimeMapper : ICellValueMapper
{
    private string[] _formats = new string[] { "G" };

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

            _formats = value;
        }
    }

    /// <summary>
    /// Gets or sets the IFormatProvider used to map the value to a DateTime.
    /// </summary>
    public IFormatProvider Provider { get; set; }

    /// <summary>
    /// Gets or sets the DateTimeStyles used to map the value to a DateTime.
    /// </summary>
    public DateTimeStyles Style { get; set; }

    public CellValueMapperResult MapCell(ExcelCell cell, CellValueMapperResult previous, MemberInfo member)
    {
        if (previous.Value is DateTime dateTimeValue)
        {
            return CellValueMapperResult.Success(dateTimeValue);
        }

        string stringValue = previous.Value?.ToString();
        try
        {
            DateTime result = DateTime.ParseExact(stringValue, Formats, Provider, Style);
            return CellValueMapperResult.Success(result);
        }
        catch (Exception exception)
        {
            return CellValueMapperResult.Invalid(exception);
        }
    }
}

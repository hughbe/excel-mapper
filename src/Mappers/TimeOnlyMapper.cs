#if NET6_0_OR_GREATER
using System.Globalization;

namespace ExcelMapper.Mappers;

/// <summary>
/// A mapper that tries to map the value of a cell to a TimeOnly.
/// </summary>
public class TimeOnlyMapper : ICellMapper, IFormatsCellMapper, ITypedCellMapper<TimeOnly>, ITypedCellMapper<TimeOnly?>
{
    private string[] _formats = ["t"];

    /// <summary>
    /// Gets or sets the date formats used to map the value to a TimeOnly.
    /// This defaults to "t" - the default TimeOnly format.
    /// While the formats haven't been set, values that don't match also try the culture's long time ("T") and 24-hour times such as "14:30" or "14:30:15".
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

        // A number in a cell without a time format is a fraction of a day, for example 0.25 for 06:00.
        if (readResult.GetValue() is double fractionOfDay && fractionOfDay >= 0 && fractionOfDay < 1)
        {
            return CellMapperResult.Success(new TimeOnly((long)Math.Round(fractionOfDay * TimeSpan.TicksPerDay)));
        }
        
        var stringValue = readResult.GetString();
        var formats = Formats;
        var provider = Provider;
        var style = Style;
        try
        {
            if (TimeOnly.TryParseExact(stringValue, formats, provider, style, out var result) ||
                (!_formatsSet && FormatUtilities.HasDigit(stringValue) && TimeOnly.TryParseExact(stringValue, FormatUtilities.GetDefaultTimeOnlyFormats(provider), provider, style, out result)))
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

    CellConversion ITypedCellMapper<TimeOnly>.Convert(ReadCellResult cell, out TimeOnly value)
    {
        var cellValue = cell.GetValue();
        if (cellValue is TimeSpan timeSpanValue)
        {
            try
            {
                value = TimeOnly.FromTimeSpan(timeSpanValue);
                return CellConversion.Success;
            }
            catch
            {
                value = default;
                return CellConversion.Invalid;
            }
        }

        if (cellValue is DateTime dateTimeValue)
        {
            value = TimeOnly.FromDateTime(dateTimeValue);
            return CellConversion.Success;
        }

        if (cellValue is double fractionOfDay && fractionOfDay >= 0 && fractionOfDay < 1)
        {
            value = new TimeOnly((long)Math.Round(fractionOfDay * TimeSpan.TicksPerDay));
            return CellConversion.Success;
        }

        var stringValue = cell.GetString();
        var provider = Provider;
        var style = Style;
        try
        {
            return TimeOnly.TryParseExact(stringValue, Formats, provider, style, out value) ||
                (!_formatsSet && FormatUtilities.HasDigit(stringValue) && TimeOnly.TryParseExact(stringValue, FormatUtilities.GetDefaultTimeOnlyFormats(provider), provider, style, out value))
                ? CellConversion.Success
                : CellConversion.Invalid;
        }
        catch
        {
            value = default;
            return CellConversion.Invalid;
        }
    }

    CellConversion ITypedCellMapper<TimeOnly?>.Convert(ReadCellResult cell, out TimeOnly? value)
    {
        var conversion = ((ITypedCellMapper<TimeOnly>)this).Convert(cell, out var result);
        value = conversion == CellConversion.Invalid ? null : result;
        return conversion;
    }

    // A separate method so that the closure capturing these values is only allocated for invalid
    // values. Creating the lambda in Map would allocate the closure for every cell.
    private static CellMapperResult InvalidResult(string? stringValue, string[] formats, IFormatProvider? provider, DateTimeStyles style)
        => MapperUtilities.InvalidDate(stringValue, style, "TimeOnly", formats, () => TimeOnly.ParseExact(stringValue!, formats, provider, style));
}
#endif

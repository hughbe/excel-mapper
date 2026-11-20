#if NET7_0_OR_GREATER
using System.Globalization;
using System.Numerics;

namespace ExcelMapper.Mappers;

/// <summary>
/// A mapper that tries to map the value of a cell to a number base type using the
/// <see cref="INumberBase{T}"/> interface.
/// </summary>
public class INumberBaseMapper<T> : INumberStyleCellMapper where T : INumberBase<T>
{
    /// <summary>
    /// Gets or sets the number styles used when mapping the value of a cell to a number base type.
    /// </summary>
    public NumberStyles Style { get; set; } = NumberStyles.Number;

    /// <inheritdoc/>
    public IFormatProvider? Provider { get; set; }

    /// <inheritdoc/>
    public CellMapperResult Map(ReadCellResult readResult)
    {
        // Try to get the value as a double first for performance.
        var value = readResult.GetValue();
        if (value is double doubleValue)
        {
            try
            {
                return CellMapperResult.Success(T.CreateChecked(doubleValue));
            }
            catch (Exception exception)
            {
                return CellMapperResult.Invalid(exception);
            }
        }

        // Fallback to parsing from string.
        var stringValue = readResult.GetString();
        var style = Style;
        var provider = Provider;

        try
        {
            if (T.TryParse(GetNumberSpan(stringValue, style), style, provider, out var result))
            {
                return CellMapperResult.Success(result);
            }
        }
        catch (Exception exception)
        {
            // TryParse throws for invalid styles.
            return CellMapperResult.Invalid(exception);
        }

        // Only create the exception (by parsing again) if it is needed.
        return MapperUtilities.InvalidFromParse(() => T.Parse(GetNumberSpan(stringValue, style), style, provider));
    }

    private static ReadOnlySpan<char> GetNumberSpan(string? stringValue, NumberStyles style)
    {
        var span = stringValue.AsSpan();

        // If this is hex, remove a 0x or &H prefix if present.
        if ((style & NumberStyles.AllowHexSpecifier) != 0 &&
            (span.StartsWith("0x", StringComparison.OrdinalIgnoreCase) || span.StartsWith("&H", StringComparison.OrdinalIgnoreCase)))
        {
            span = span[2..];
        }

        return span;
    }
}
#endif

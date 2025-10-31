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
        var stringValue = readResult.GetString().AsSpan();

        // If this is hex.
        if (Style.HasFlag(NumberStyles.AllowHexSpecifier))
        {
            // Remove 0x or &H prefix if present.
            if (stringValue.StartsWith("0x", StringComparison.OrdinalIgnoreCase) ||
                stringValue.StartsWith("&H", StringComparison.OrdinalIgnoreCase))
            {
                stringValue = stringValue[2..];
            }
        }

        try
        {
            var result = T.Parse(stringValue, Style, Provider);
            return CellMapperResult.Success(result);
        }
        catch (Exception exception)
        {
            return CellMapperResult.Invalid(exception);
        }
    }
}

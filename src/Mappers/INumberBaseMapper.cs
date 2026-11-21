using System.Diagnostics.CodeAnalysis;
﻿#if NET7_0_OR_GREATER
using System.Globalization;
using System.Numerics;

namespace ExcelMapper.Mappers;

/// <summary>
/// A mapper that tries to map the value of a cell to a number base type using the
/// <see cref="INumberBase{T}"/> interface.
/// </summary>
public class INumberBaseMapper<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.Interfaces)] T> : INumberStyleCellMapper, ITypedCellMapper<T> where T : INumberBase<T>
{
    /// <summary>
    /// Gets or sets the number styles used when mapping the value of a cell to a number base type.
    /// </summary>
    /// <remarks>
    /// The default is <see cref="NumberStyles.Number"/>, plus <see cref="NumberStyles.AllowExponent"/>
    /// for floating-point types so that values such as "1E5" can be parsed.
    /// </remarks>
    public NumberStyles Style { get; set; } = s_defaultStyle;

    private static readonly NumberStyles s_defaultStyle = typeof(T).ImplementsGenericInterface(typeof(IFloatingPoint<>), out _)
        ? NumberStyles.Number | NumberStyles.AllowExponent
        : NumberStyles.Number;

    private static readonly bool s_isInteger = IsIntegerType(typeof(T));

    private static bool IsIntegerType([DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.Interfaces)] Type type)
    {
        foreach (var interfaceType in type.GetInterfaces())
        {
            if (interfaceType.IsGenericType && interfaceType.GetGenericTypeDefinition() == typeof(IBinaryInteger<>))
            {
                return true;
            }
        }

        return false;
    }

    /// <inheritdoc/>
    public IFormatProvider? Provider { get; set; }

    /// <inheritdoc/>
    public CellMapperResult Map(ReadCellResult readResult)
    {
        // Try to get the value as a double first for performance.
        var value = readResult.GetValue();
        if (value is double doubleValue)
        {
            // Excel stores every number as a double. Don't silently drop the fractional part when
            // mapping to an integer type: "2.7" as text can't be parsed as an int, so neither can 2.7.
            if (s_isInteger && !double.IsInteger(doubleValue) && double.IsFinite(doubleValue))
            {
                return NotWholeNumberResult(doubleValue);
            }

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
        return InvalidResult(stringValue, style, provider);
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

    // A separate method so that the closure capturing these values is only allocated for invalid
    // values. Creating the lambda in Map would allocate the closure for every cell.
    CellConversion ITypedCellMapper<T>.Convert(ReadCellResult cell, out T value)
    {
        if (cell.GetValue() is double doubleValue)
        {
            if (s_isInteger && !double.IsInteger(doubleValue) && double.IsFinite(doubleValue))
            {
                value = default!;
                return CellConversion.Invalid;
            }

            try
            {
                value = T.CreateChecked(doubleValue);
                return CellConversion.Success;
            }
            catch
            {
                value = default!;
                return CellConversion.Invalid;
            }
        }

        var stringValue = cell.GetString();
        var style = Style;
        try
        {
            return T.TryParse(GetNumberSpan(stringValue, style), style, Provider, out value!) ? CellConversion.Success : CellConversion.Invalid;
        }
        catch
        {
            value = default!;
            return CellConversion.Invalid;
        }
    }

    private static CellMapperResult NotWholeNumberResult(double value)
        => CellMapperResult.Invalid(() => new FormatException($"The value {value} is not a whole number, so it can't be converted to {typeof(T)}."));

    // A separate method so that the closure capturing these values is only allocated for invalid
    // values. Creating the lambda in Map would allocate the closure for every cell.
    private static CellMapperResult InvalidResult(string? stringValue, NumberStyles style, IFormatProvider? provider)
        => MapperUtilities.InvalidFromParse(() => T.Parse(GetNumberSpan(stringValue, style), style, provider));
}
#endif

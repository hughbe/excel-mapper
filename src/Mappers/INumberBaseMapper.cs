using System.Diagnostics.CodeAnalysis;
﻿#if NET7_0_OR_GREATER
using System.Globalization;
using System.Numerics;

namespace ExcelMapper.Mappers;

/// <summary>
/// A mapper that tries to map the value of a cell to a number base type using the
/// <see cref="INumberBase{T}"/> interface.
/// </summary>
internal class INumberBaseMapper<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.Interfaces)] T> : INumberStyleCellMapper, ICellMapper<T> where T : INumberBase<T>
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

    CellConversion ICellMapper<T>.Convert(ReadCellResult cell, in CellContext context, out T value, ref object? error)
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
            return T.TryParse(GetNumberSpan(stringValue, style), style, Provider ?? cell.Culture, out value!) ? CellConversion.Success : CellConversion.Invalid;
        }
        catch
        {
            value = default!;
            return CellConversion.Invalid;
        }
    }

    public Exception CreateException(ReadCellResult cell, object? error)
    {
        if (cell.GetValue() is double doubleValue)
        {
            // Excel stores every number as a double. A fractional number isn't silently truncated
            // for an integer type: "2.7" as text can't be parsed as an int, so neither can 2.7.
            return s_isInteger && !double.IsInteger(doubleValue) && double.IsFinite(doubleValue)
                ? MapperUtilities.NotWholeNumber(doubleValue, typeof(T))
                : MapperUtilities.InvalidFromParse(() => T.CreateChecked(doubleValue));
        }

        var stringValue = cell.GetString();
        var style = Style;
        var provider = Provider ?? cell.Culture;
        try
        {
            // TryParse throws for invalid styles.
            _ = T.TryParse(GetNumberSpan(stringValue, style), style, provider, out _);
        }
        catch (Exception exception)
        {
            return exception;
        }

        return MapperUtilities.InvalidNumber(stringValue, style, provider, typeof(T), s_isInteger, s_isBuiltIn, () => T.Parse(GetNumberSpan(stringValue, style), style, provider));
    }

    // Whether T is a built-in number type, whose invalid values' exceptions can be created without
    // parsing again (see MapperUtilities.InvalidNumber).
    private static readonly bool s_isBuiltIn = Type.GetTypeCode(typeof(T)) is >= TypeCode.SByte and <= TypeCode.Decimal || typeof(T) == typeof(Int128) || typeof(T) == typeof(UInt128) || typeof(T) == typeof(Half);
}
#endif

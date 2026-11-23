#if NET7_0_OR_GREATER
using System.Globalization;
using System.Numerics;

namespace ExcelMapper.Mappers;

/// <summary>
/// A mapper that tries to map the value of a cell to a number base type using the
/// <see cref="INumberBase{T}"/> interface.
/// </summary>
internal class INumberBaseMapper<T> : ICellMapper<T> where T : INumberBase<T>
{
    // Whether T only has whole numbers, so a fractional number is invalid rather than truncated.
    // Worked out with generic math rather than by looking for IBinaryInteger<T> among T's
    // interfaces, which trimming could remove: a type that can't hold 0.5 has no fractions.
    private static readonly bool s_isInteger = HasOnlyWholeNumbers();

    // Types with fractions accept exponents ("1E5"), as double.Parse does. Complex numbers don't,
    // as Complex.Parse doesn't by default.
    private static readonly NumberStyles s_defaultStyle = s_isInteger || typeof(T) == typeof(Complex)
        ? NumberStyles.Number
        : NumberStyles.Number | NumberStyles.AllowExponent;

    private static bool HasOnlyWholeNumbers()
    {
        try
        {
            return T.IsInteger(T.CreateTruncating(0.5));
        }
        catch
        {
            // A type that can't be created from a double.
            return false;
        }
    }

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

            // Out of range values are checked rather than caught, as an exception for every out of
            // range cell is slow. Only other number types (a user's) need catching.
            if (s_range != NumberRange.Unknown)
            {
                if (!IsInRange(doubleValue))
                {
                    value = default!;
                    return CellConversion.Invalid;
                }

                value = T.CreateChecked(doubleValue);
                return CellConversion.Success;
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
        var style = context.Options.NumberStyle ?? s_defaultStyle;
        var provider = context.Options.Provider ?? cell.Culture;
        try
        {
            // Most cells of whole numbers are plain digits, which parse much faster as an integer.
            // Anything else, such as "1,000", is parsed with the member's style.
            var span = GetNumberSpan(stringValue, style);
            if (s_hasFastIntegerParse && MapperUtilities.AllowsInteger(style) && T.TryParse(span, NumberStyles.Integer, provider, out value!))
            {
                return CellConversion.Success;
            }

            return T.TryParse(span, style, provider, out value!) ? CellConversion.Success : CellConversion.Invalid;
        }
        catch
        {
            value = default!;
            return CellConversion.Invalid;
        }
    }

    public Exception CreateException(ReadCellResult cell, in CellContext context, object? error)
    {
        if (cell.GetValue() is double doubleValue)
        {
            // Excel stores every number as a double. A fractional number isn't silently truncated
            // for an integer type: "2.7" as text can't be parsed as an int, so neither can 2.7.
            return s_isInteger && !double.IsInteger(doubleValue) && double.IsFinite(doubleValue)
                ? MapperUtilities.NotWholeNumber(doubleValue, typeof(T))
                : s_range != NumberRange.Unknown && !IsInRange(doubleValue)
                ? MapperUtilities.OutOfRange(doubleValue, typeof(T))
                : MapperUtilities.InvalidFromParse(() => T.CreateChecked(doubleValue));
        }

        var stringValue = cell.GetString();
        var style = context.Options.NumberStyle ?? s_defaultStyle;
        var provider = context.Options.Provider ?? cell.Culture;
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

    // Which doubles T.CreateChecked converts without an OverflowException, for the built-in number
    // types: Unknown for other types, whose conversions are caught.
    private enum NumberRange
    {
        Unknown,
        Any,
        Finite,
        TypeCode,
        Int128,
        UInt128,
    }

    private static readonly NumberRange s_range = GetRange();

    private static NumberRange GetRange()
    {
        var type = typeof(T);
        return Type.GetTypeCode(type) is >= TypeCode.SByte and <= TypeCode.Decimal ? NumberRange.TypeCode
            : type == typeof(Half) || type == typeof(Complex) ? NumberRange.Any
            : type == typeof(BigInteger) ? NumberRange.Finite
            : type == typeof(Int128) ? NumberRange.Int128
            : type == typeof(UInt128) ? NumberRange.UInt128
            : type == typeof(nint) || type == typeof(nuint) ? NumberRange.TypeCode
            : NumberRange.Unknown;
    }

    // nint and nuint have the range of the int types of the process's pointer size.
    private static readonly TypeCode s_typeCode =
        typeof(T) == typeof(nint) ? (IntPtr.Size == 8 ? TypeCode.Int64 : TypeCode.Int32)
        : typeof(T) == typeof(nuint) ? (IntPtr.Size == 8 ? TypeCode.UInt64 : TypeCode.UInt32)
        : Type.GetTypeCode(typeof(T));

    // 2^127 and 2^128 are the first values out of range, and are exact as doubles.
    private const double TwoTo127 = 170141183460469231731687303715884105728.0;
    private const double TwoTo128 = 340282366920938463463374607431768211456.0;

    private static bool IsInRange(double value) => s_range switch
    {
        NumberRange.TypeCode => MapperUtilities.IsInRange(value, s_typeCode),
        NumberRange.Finite => double.IsFinite(value),
        NumberRange.Int128 => value >= -TwoTo127 && value < TwoTo127,
        NumberRange.UInt128 => value >= 0 && value < TwoTo128,
        _ => true,
    };

    // Whether T is a built-in integer type, which parses NumberStyles.Integer on a fast path.
    // Other number types (such as BigInteger, or a user's) aren't parsed twice.
    private static readonly bool s_hasFastIntegerParse = Type.GetTypeCode(typeof(T)) is >= TypeCode.SByte and <= TypeCode.UInt64 || typeof(T) == typeof(Int128) || typeof(T) == typeof(UInt128);

    // Whether T is a built-in number type, whose invalid values' exceptions can be created without
    // parsing again (see MapperUtilities.InvalidNumber).
    private static readonly bool s_isBuiltIn = Type.GetTypeCode(typeof(T)) is >= TypeCode.SByte and <= TypeCode.Decimal || typeof(T) == typeof(Int128) || typeof(T) == typeof(UInt128) || typeof(T) == typeof(Half);
}
#endif

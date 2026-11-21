#if !NET7_0_OR_GREATER
using System.Globalization;

namespace ExcelMapper.Mappers;

/// <summary>
/// Maps cells to the built-in numeric types on the netstandard builds, which don't have
/// <c>INumberBase&lt;T&gt;</c>. It behaves like <c>INumberBaseMapper&lt;T&gt;</c> in the net8.0 build:
/// the same default number styles, support for <see cref="ExcelNumberStyleAttribute"/> and a format
/// provider, and the same rules for typed number cells.
/// </summary>
internal sealed class NumberMapper : INumberStyleCellMapper
{
    private readonly TypeCode _typeCode;

    public NumberMapper(Type type)
    {
        ThrowHelpers.ThrowIfNull(type, nameof(type));
        if (!IsSupportedType(type))
        {
            throw new ArgumentException($"Type \"{type}\" is not a built-in numeric type.", nameof(type));
        }

        Type = type;
        _typeCode = Type.GetTypeCode(type);
        Style = GetDefaultStyle(_typeCode);
    }

    public Type Type { get; }

    public NumberStyles Style { get; set; }

    public IFormatProvider? Provider { get; set; }

    /// <summary>
    /// Gets whether the type is one of the built-in numeric types this mapper supports.
    /// </summary>
    public static bool IsSupportedType(Type type) => !type.IsEnum && IsNumeric(Type.GetTypeCode(type));

    private static bool IsNumeric(TypeCode typeCode) => typeCode is >= TypeCode.SByte and <= TypeCode.Decimal;

    private static bool IsInteger(TypeCode typeCode) => typeCode is >= TypeCode.SByte and <= TypeCode.UInt64;

    /// <summary>
    /// The default number styles, the same as the net8.0 build: <see cref="NumberStyles.Number"/>,
    /// plus exponents for floating-point types.
    /// </summary>
    internal static NumberStyles GetDefaultStyle(TypeCode typeCode)
        => IsInteger(typeCode) ? NumberStyles.Number : NumberStyles.Number | NumberStyles.AllowExponent;

    public CellMapperResult Map(ReadCellResult readResult)
    {
        var value = readResult.GetValue();

        // Excel stores every number as a double.
        if (value is double doubleValue)
        {
            // Pass the boxed value on, rather than boxing the double again.
            return FromDouble(doubleValue, value);
        }

        // Other typed values, such as booleans and dates, are converted as before.
        if (value is not null and not string)
        {
            try
            {
                return CellMapperResult.Success(Convert.ChangeType(value, Type, CultureInfo.InvariantCulture));
            }
            catch (Exception exception)
            {
                return CellMapperResult.Invalid(exception);
            }
        }

        var stringValue = GetNumberText(readResult.GetString(), Style);
        var style = Style;
        var provider = Provider;
        try
        {
            if (TryParse(_typeCode, stringValue, style, provider, out var result))
            {
                return CellMapperResult.Success(result);
            }
        }
        catch (Exception exception)
        {
            // TryParse throws for invalid styles, such as hex for floating-point types.
            return CellMapperResult.Invalid(exception);
        }

        // Only create the exception (by parsing again) if it is needed.
        return InvalidResult(_typeCode, stringValue, style, provider);
    }

    private CellMapperResult FromDouble(double value, object boxedValue)
    {
        // Don't silently drop the fractional part when mapping to an integer type: "2.7" as text
        // can't be parsed as an int, so neither can 2.7.
        if (IsInteger(_typeCode) && Math.Floor(value) != value && !double.IsNaN(value) && !double.IsInfinity(value))
        {
            return NotWholeNumberResult(value, Type);
        }

        try
        {
            // A double cell mapped to a double needs no conversion.
            return CellMapperResult.Success(_typeCode == TypeCode.Double ? boxedValue : Convert.ChangeType(boxedValue, Type, CultureInfo.InvariantCulture));
        }
        catch (Exception exception)
        {
            // For example, out of range for the type.
            return CellMapperResult.Invalid(exception);
        }
    }

    private static string GetNumberText(string? stringValue, NumberStyles style)
    {
        var text = stringValue ?? string.Empty;

        // If this is hex, remove a 0x or &H prefix if present.
        if ((style & NumberStyles.AllowHexSpecifier) != 0 &&
            (text.StartsWith("0x", StringComparison.OrdinalIgnoreCase) || text.StartsWith("&H", StringComparison.OrdinalIgnoreCase)))
        {
            text = text.Substring(2);
        }

        return text;
    }

    private static bool TryParse(TypeCode typeCode, string text, NumberStyles style, IFormatProvider? provider, out object? result)
    {
        bool success;
        switch (typeCode)
        {
            case TypeCode.SByte: success = sbyte.TryParse(text, style, provider, out var sbyteValue); result = sbyteValue; break;
            case TypeCode.Byte: success = byte.TryParse(text, style, provider, out var byteValue); result = byteValue; break;
            case TypeCode.Int16: success = short.TryParse(text, style, provider, out var shortValue); result = shortValue; break;
            case TypeCode.UInt16: success = ushort.TryParse(text, style, provider, out var ushortValue); result = ushortValue; break;
            case TypeCode.Int32: success = int.TryParse(text, style, provider, out var intValue); result = intValue; break;
            case TypeCode.UInt32: success = uint.TryParse(text, style, provider, out var uintValue); result = uintValue; break;
            case TypeCode.Int64: success = long.TryParse(text, style, provider, out var longValue); result = longValue; break;
            case TypeCode.UInt64: success = ulong.TryParse(text, style, provider, out var ulongValue); result = ulongValue; break;
            case TypeCode.Single: success = float.TryParse(text, style, provider, out var floatValue); result = floatValue; break;
            case TypeCode.Double: success = double.TryParse(text, style, provider, out var doubleValue); result = doubleValue; break;
            case TypeCode.Decimal: success = decimal.TryParse(text, style, provider, out var decimalValue); result = decimalValue; break;
            default: throw new InvalidOperationException($"Unexpected type code {typeCode}.");
        }

        return success;
    }

    private static object Parse(TypeCode typeCode, string text, NumberStyles style, IFormatProvider? provider) => typeCode switch
    {
        TypeCode.SByte => sbyte.Parse(text, style, provider),
        TypeCode.Byte => byte.Parse(text, style, provider),
        TypeCode.Int16 => short.Parse(text, style, provider),
        TypeCode.UInt16 => ushort.Parse(text, style, provider),
        TypeCode.Int32 => int.Parse(text, style, provider),
        TypeCode.UInt32 => uint.Parse(text, style, provider),
        TypeCode.Int64 => long.Parse(text, style, provider),
        TypeCode.UInt64 => ulong.Parse(text, style, provider),
        TypeCode.Single => float.Parse(text, style, provider),
        TypeCode.Double => double.Parse(text, style, provider),
        TypeCode.Decimal => decimal.Parse(text, style, provider),
        _ => throw new InvalidOperationException($"Unexpected type code {typeCode}.")
    };

    // Separate methods so that closures capturing these values are only allocated for invalid
    // values. Creating the lambdas in Map would allocate a closure for every cell.
    private static CellMapperResult InvalidResult(TypeCode typeCode, string text, NumberStyles style, IFormatProvider? provider)
        => MapperUtilities.InvalidFromParse(() => Parse(typeCode, text, style, provider));

    private static CellMapperResult NotWholeNumberResult(double value, Type type)
        => CellMapperResult.Invalid(() => new FormatException($"The value {value} is not a whole number, so it can't be converted to {type}."));
}
#endif

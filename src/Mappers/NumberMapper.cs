#if !NET7_0_OR_GREATER
using System.Globalization;

namespace ExcelMapper.Mappers;

/// <summary>
/// Maps cells to the built-in numeric types on the netstandard builds, which don't have
/// <c>INumberBase&lt;T&gt;</c>. It behaves like <c>INumberBaseMapper&lt;T&gt;</c> in the net8.0 build:
/// the same default number styles, support for <see cref="ExcelNumberStyleAttribute"/> and a format
/// provider, and the same rules for typed number cells. Created with <see cref="Create"/>, which
/// returns a <see cref="NumberMapper{T}"/> that also converts cells without boxing.
/// </summary>
internal abstract class NumberMapper : INumberStyleCellMapper
{
    private readonly TypeCode _typeCode;

    /// <summary>
    /// Creates the mapper for a built-in numeric type.
    /// </summary>
    public static NumberMapper Create(Type type)
    {
        ThrowHelpers.ThrowIfNull(type, nameof(type));
        if (!IsSupportedType(type))
        {
            throw new ArgumentException($"Type \"{type}\" is not a built-in numeric type.", nameof(type));
        }

        return Type.GetTypeCode(type) switch
        {
            TypeCode.SByte => new NumberMapper<sbyte>(),
            TypeCode.Byte => new NumberMapper<byte>(),
            TypeCode.Int16 => new NumberMapper<short>(),
            TypeCode.UInt16 => new NumberMapper<ushort>(),
            TypeCode.Int32 => new NumberMapper<int>(),
            TypeCode.UInt32 => new NumberMapper<uint>(),
            TypeCode.Int64 => new NumberMapper<long>(),
            TypeCode.UInt64 => new NumberMapper<ulong>(),
            TypeCode.Single => new NumberMapper<float>(),
            TypeCode.Double => new NumberMapper<double>(),
            _ => new NumberMapper<decimal>(),
        };
    }

    private protected NumberMapper(Type type)
    {
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

    private protected static bool IsInteger(TypeCode typeCode) => typeCode is >= TypeCode.SByte and <= TypeCode.UInt64;

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

    private protected TypeCode TypeCode => _typeCode;

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

    private protected static string GetNumberText(string? stringValue, NumberStyles style)
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

/// <summary>
/// A <see cref="NumberMapper"/> for <typeparamref name="T"/> that converts cells without boxing,
/// making the same decisions as <see cref="NumberMapper.Map"/>.
/// </summary>
internal sealed class NumberMapper<T> : NumberMapper, ICellConverter<T>, ICellConverter<T?> where T : struct
{
    private delegate bool TryParser(string text, NumberStyles style, IFormatProvider? provider, out T result);

    // Created once for each type, so that parsing and converting doubles don't box.
    private static readonly TryParser s_tryParse = CreateTryParse();
    private static readonly Func<double, T> s_fromDouble = CreateFromDouble();

    public NumberMapper() : base(typeof(T))
    {
    }

    private static TryParser CreateTryParse()
    {
        object tryParse = Type.GetTypeCode(typeof(T)) switch
        {
            TypeCode.SByte => new NumberMapper<sbyte>.TryParser(sbyte.TryParse),
            TypeCode.Byte => new NumberMapper<byte>.TryParser(byte.TryParse),
            TypeCode.Int16 => new NumberMapper<short>.TryParser(short.TryParse),
            TypeCode.UInt16 => new NumberMapper<ushort>.TryParser(ushort.TryParse),
            TypeCode.Int32 => new NumberMapper<int>.TryParser(int.TryParse),
            TypeCode.UInt32 => new NumberMapper<uint>.TryParser(uint.TryParse),
            TypeCode.Int64 => new NumberMapper<long>.TryParser(long.TryParse),
            TypeCode.UInt64 => new NumberMapper<ulong>.TryParser(ulong.TryParse),
            TypeCode.Single => new NumberMapper<float>.TryParser(float.TryParse),
            TypeCode.Double => new NumberMapper<double>.TryParser(double.TryParse),
            _ => new NumberMapper<decimal>.TryParser(decimal.TryParse),
        };
        return (TryParser)tryParse;
    }

    // The same results as Convert.ChangeType for whole numbers (fractions are rejected first): out
    // of range, NaN and infinity throw OverflowException.
    private static Func<double, T> CreateFromDouble()
    {
        object fromDouble = Type.GetTypeCode(typeof(T)) switch
        {
            TypeCode.SByte => new Func<double, sbyte>(d => checked((sbyte)d)),
            TypeCode.Byte => new Func<double, byte>(d => checked((byte)d)),
            TypeCode.Int16 => new Func<double, short>(d => checked((short)d)),
            TypeCode.UInt16 => new Func<double, ushort>(d => checked((ushort)d)),
            TypeCode.Int32 => new Func<double, int>(d => checked((int)d)),
            TypeCode.UInt32 => new Func<double, uint>(d => checked((uint)d)),
            TypeCode.Int64 => new Func<double, long>(d => checked((long)d)),
            TypeCode.UInt64 => new Func<double, ulong>(d => checked((ulong)d)),
            TypeCode.Single => new Func<double, float>(d => (float)d),
            TypeCode.Double => new Func<double, double>(d => d),
            _ => new Func<double, decimal>(d => (decimal)d),
        };
        return (Func<double, T>)fromDouble;
    }

    CellConversion ICellConverter<T>.Convert(ReadCellResult cell, out T value)
    {
        var cellValue = cell.GetValue();
        try
        {
            if (cellValue is double doubleValue)
            {
                if (IsInteger(TypeCode) && Math.Floor(doubleValue) != doubleValue && !double.IsNaN(doubleValue) && !double.IsInfinity(doubleValue))
                {
                    value = default;
                    return CellConversion.Invalid;
                }

                value = s_fromDouble(doubleValue);
                return CellConversion.Success;
            }

            // Other typed values, such as booleans and dates, are rare: convert them as Map does.
            if (cellValue is not null and not string)
            {
                value = (T)System.Convert.ChangeType(cellValue, typeof(T), CultureInfo.InvariantCulture);
                return CellConversion.Success;
            }

            return s_tryParse(GetNumberText(cell.GetString(), Style), Style, Provider, out value)
                ? CellConversion.Success
                : CellConversion.Invalid;
        }
        catch
        {
            // For example, out of range, or an invalid style such as hex for floating-point types.
            value = default;
            return CellConversion.Invalid;
        }
    }

    CellConversion ICellConverter<T?>.Convert(ReadCellResult cell, out T? value)
    {
        var conversion = ((ICellConverter<T>)this).Convert(cell, out var result);
        value = conversion == CellConversion.Invalid ? null : result;
        return conversion;
    }
}
#endif

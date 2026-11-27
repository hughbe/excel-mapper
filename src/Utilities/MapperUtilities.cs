using System.Globalization;

namespace ExcelMapper.Utilities;

internal static class MapperUtilities
{
    /// <summary>
    /// Whether text that parses with <see cref="NumberStyles.Integer"/> parses to the same whole
    /// number with <paramref name="style"/>: the style allows everything Integer does (leading and
    /// trailing white space and a leading sign), and isn't hexadecimal or binary. Then a whole number
    /// can be parsed with Integer first, which .NET parses several times faster than styles that
    /// allow thousands separators or a decimal point, such as the default
    /// <see cref="NumberStyles.Number"/>.
    /// </summary>
    public static bool AllowsInteger(NumberStyles style)
        => (style & NumberStyles.Integer) == NumberStyles.Integer && (style & NonDecimalStyles) == 0;

#if NET8_0_OR_GREATER
    private const NumberStyles NonDecimalStyles = NumberStyles.AllowHexSpecifier | NumberStyles.AllowBinarySpecifier;
#else
    private const NumberStyles NonDecimalStyles = NumberStyles.AllowHexSpecifier;
#endif

    /// <summary>
    /// Gets the exception for a value that couldn't be parsed by running the throwing version of
    /// the parse that failed. This keeps the exact exception (for example FormatException or
    /// OverflowException), but throwing and catching it is slow, so it's only used where the
    /// exception can't be created directly (see the other methods). Like all of these methods, it's
    /// only called from <see cref="ICellExceptionSource.CreateException"/>, when the exception is used.
    /// </summary>
    public static Exception InvalidFromParse(Action parse)
    {
        try
        {
            parse();
        }
        catch (Exception exception)
        {
            return exception;
        }

        return new FormatException("The value could not be parsed.");
    }

    /// <summary>
    /// Creates a <see cref="FormatException"/> without throwing it: for parsers whose only failure for text is a
    /// <see cref="FormatException"/>, such as <see cref="Guid.Parse(string)"/>. Null text (which
    /// the parser rejects with an <see cref="ArgumentNullException"/>) uses <paramref name="parse"/>.
    /// </summary>
    public static Exception InvalidFormat(string? text, string description, Action parse)
        => text is null
            ? InvalidFromParse(parse)
            : new FormatException($"The value \"{text}\" isn't {description}.");

    /// <summary>
    /// Creates the exception for a date or time that couldn't be parsed with
    /// <paramref name="formats"/>: a <see cref="FormatException"/>, created without throwing it. Styles other than whitespace and UTC can make ParseExact throw an
    /// <see cref="ArgumentException"/> instead, so they use <paramref name="parse"/>.
    /// </summary>
    public static Exception InvalidDate(string? text, DateTimeStyles style, string typeName, string[] formats, Action parse)
        => (style & ~(DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal)) != 0
            ? InvalidFromParse(parse)
            : InvalidFormat(text, FormatUtilities.DescribeFormats(typeName, formats), parse);

    /// <summary>
    /// Creates the exception for a number that couldn't be parsed: the exception the type's Parse
    /// method throws, created without throwing it: an <see cref="OverflowException"/> if the text is a number that is out of range
    /// (or isn't whole, for an integer type), and otherwise a <see cref="FormatException"/>. Where
    /// that can't be told without parsing, <paramref name="parse"/> is used.
    /// </summary>
    /// <param name="text">The text that isn't a valid number.</param>
    /// <param name="style">The number style it was parsed with.</param>
    /// <param name="provider">The format provider it was parsed with.</param>
    /// <param name="type">The number type.</param>
    /// <param name="isInteger">Whether the type is an integer type.</param>
    /// <param name="canCreateException">Whether the type is a built-in number type, whose Parse
    /// method fails as described.</param>
    /// <param name="parse">Parses the text, to get the exception where it can't be created directly.</param>
    public static Exception InvalidNumber(string? text, NumberStyles style, IFormatProvider? provider, Type type, bool isInteger, bool canCreateException, Action parse)
        => !canCreateException || text is null || GetNumberFailure(text, style, provider, isInteger) is not { } failure
            ? InvalidFromParse(parse)
            : failure switch
            {
                NumberFailure.NotWhole => new OverflowException($"The value \"{text}\" isn't a whole number, so it can't be {TypeNames.Get(type)}."),
                NumberFailure.OutOfRange => new OverflowException($"The value \"{text}\" is too large or too small for {TypeNames.Get(type)}."),
                _ => new FormatException($"The value \"{text}\" isn't a valid {TypeNames.Get(type)}."),
            };

    private enum NumberFailure
    {
        NotANumber,
        OutOfRange,
        NotWhole,
    }

    /// <summary>
    /// Whether a number cell's value can be converted to a built-in number type without
    /// overflowing, so that it's checked rather than caught: converting a double to an integer
    /// type or <see cref="decimal"/> throws an <see cref="OverflowException"/> if it's out of range
    /// (or NaN or infinite), and to <see cref="float"/> or <see cref="double"/> never does.
    /// Fractions are rejected before this for integer types.
    /// </summary>
    public static bool IsInRange(double value, TypeCode typeCode) => typeCode switch
    {
        TypeCode.SByte => value >= sbyte.MinValue && value <= sbyte.MaxValue,
        TypeCode.Byte => value >= byte.MinValue && value <= byte.MaxValue,
        TypeCode.Int16 => value >= short.MinValue && value <= short.MaxValue,
        TypeCode.UInt16 => value >= ushort.MinValue && value <= ushort.MaxValue,
        TypeCode.Int32 => value >= int.MinValue && value <= int.MaxValue,
        TypeCode.UInt32 => value >= uint.MinValue && value <= uint.MaxValue,
        // 2^63, 2^64 and 2^96 are the first values out of range, and are exact as doubles, unlike
        // long.MaxValue, ulong.MaxValue and decimal.MaxValue.
        TypeCode.Int64 => value >= -TwoTo63 && value < TwoTo63,
        TypeCode.UInt64 => value >= 0 && value < TwoTo64,
        TypeCode.Decimal => value > -TwoTo96 && value < TwoTo96,
        _ => true,
    };

    internal const double TwoTo63 = 9223372036854775808.0;
    private const double TwoTo64 = 18446744073709551616.0;
    private const double TwoTo96 = 79228162514264337593543950336.0;

    /// <summary>
    /// The exception for a number cell out of the range of its type, created without converting
    /// it: an <see cref="OverflowException"/>, as for text out of range (see <see cref="InvalidNumber"/>),
    /// with the number written in the importer's culture, as the cell's text is.
    /// </summary>
    public static Exception OutOfRange(double value, Type type, IFormatProvider provider)
        => new OverflowException($"The value {value.ToString(provider)} is too large or too small for {TypeNames.Get(type)}.");

    /// <summary>
    /// Removes floating-point error from a number in a cell that isn't whole, if that makes it
    /// whole. Excel shows and calculates with 15 significant digits, so a formula such as
    /// <c>=1.15*100</c>, which Excel stores as the double 114.99999999999999, is 115 to the user:
    /// the digits past the 15th are rounding error, not a fraction. A number that's still
    /// fractional with 15 significant digits, such as 2.7, is returned unchanged.
    /// </summary>
    public static double RemoveFloatingPointError(double value)
    {
        var magnitude = Math.Abs(value);

        // From 1e15 the 15th significant digit is above the units, so rounding to it would change
        // the whole number. Below 0.1 rounding can't give a whole number other than zero, and a
        // tiny number such as 5.6E-17 (=0.1+0.2-0.3) is what Excel shows too. NaN and infinity
        // fail both checks.
        if (!(magnitude >= 0.1 && magnitude < 1e15))
        {
            return value;
        }

        // A real fraction, such as 2.5, is far from the nearest whole number, so skip the slower
        // rounding to 15 significant digits below. Rounding error is at most half a unit in the
        // 15th digit, which is less than 5e-15 of the number.
        var nearest = Math.Round(value);
        if (Math.Abs(value - nearest) > magnitude * 1e-14)
        {
            return value;
        }

        var decimals = 14 - (int)Math.Floor(Math.Log10(magnitude));
        var rounded = Math.Round(value, Math.Max(0, Math.Min(15, decimals)));
        return rounded == nearest ? nearest : value;
    }

    // The exception for a number in a cell that isn't whole, mapped to an integer type: an
    // OverflowException, as for text such as "2.5" (see InvalidNumber), with the number written
    // in the importer's culture, as the cell's text is.
    public static Exception NotWholeNumber(double value, Type type, IFormatProvider provider)
        => new OverflowException($"The value {value.ToString(provider)} isn't a whole number, so it can't be {TypeNames.Get(type)}.");

    // Why a built-in number type's TryParse rejected the text: a number out of range, or with a
    // fraction for an integer type, or not a number; or null if that can't be told without parsing.
    private static NumberFailure? GetNumberFailure(string text, NumberStyles style, IFormatProvider? provider, bool isInteger)
    {
        // Hex and binary numbers aren't parsed by decimal and double.
        if ((style & ~NumberStyles.Any) != 0)
        {
            return null;
        }

        // A number that decimal or double can parse, with the same style, is a number the type
        // can't hold.
        if (decimal.TryParse(text, style, provider, out var exact))
        {
            return isInteger && exact != decimal.Truncate(exact) ? NumberFailure.NotWhole : NumberFailure.OutOfRange;
        }

        if (!double.TryParse(text, style, provider, out var number))
        {
            return NumberFailure.NotANumber;
        }

        // Infinity and NaN symbols, or a number too large for a double.
        return double.IsNaN(number) || double.IsInfinity(number) ? null : NumberFailure.OutOfRange;
    }
}

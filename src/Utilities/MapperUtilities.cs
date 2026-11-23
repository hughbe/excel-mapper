using System.Globalization;

namespace ExcelMapper.Utilities;

internal static class MapperUtilities
{
    /// <summary>
    /// Gets the exception for a value that couldn't be parsed by running the throwing version of
    /// the parse that failed. This keeps the exact exception (for example FormatException or
    /// OverflowException), but throwing and catching it is slow, so it's only used where the
    /// exception can't be created directly (see the other methods). Like all of these methods, it's
    /// only called from <see cref="ICellMapper{T}.CreateException"/>, when the exception is used.
    /// </summary>
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
    /// <paramref name="formats"/>: a <see cref="FormatException"/>, created without throwing it. Styles other than whitespace can make ParseExact throw an
    /// <see cref="ArgumentException"/> instead, so they use <paramref name="parse"/>.
    /// </summary>
    public static Exception InvalidDate(string? text, DateTimeStyles style, string typeName, string[] formats, Action parse)
        => (style & ~DateTimeStyles.AllowWhiteSpaces) != 0
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

    // The exception for a number in a cell that isn't whole, mapped to an integer type.
    public static Exception NotWholeNumber(double value, Type type)
        => new FormatException($"The value {value} isn't a whole number, so it can't be {TypeNames.Get(type)}.");

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

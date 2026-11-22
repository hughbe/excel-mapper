using System.Globalization;

namespace ExcelMapper.Utilities;

internal static class MapperUtilities
{
    /// <summary>
    /// Returns an invalid result whose exception is created only if it is needed, by running the
    /// throwing version of a parse that has already failed. This keeps the exact exception (for
    /// example FormatException or OverflowException), but throwing and catching it is slow, so it's
    /// only used where the exception can't be created directly (see the other methods).
    /// </summary>
    public static CellMapperResult InvalidFromParse(Action parse)
        => CellMapperResult.Invalid(() =>
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
        });

    /// <summary>
    /// Returns an invalid result whose exception is a <see cref="FormatException"/> created only if
    /// it is needed, without throwing it: for parsers whose only failure for text is a
    /// <see cref="FormatException"/>, such as <see cref="Guid.Parse(string)"/>. Null text (which
    /// the parser rejects with an <see cref="ArgumentNullException"/>) uses <paramref name="parse"/>.
    /// </summary>
    public static CellMapperResult InvalidFormat(string? text, string description, Action parse)
        => text is null
            ? InvalidFromParse(parse)
            : CellMapperResult.Invalid(() => new FormatException($"The value \"{text}\" isn't {description}."));

    /// <summary>
    /// Returns an invalid result for a date or time that couldn't be parsed with
    /// <paramref name="formats"/>, whose exception is a <see cref="FormatException"/> created only if
    /// it is needed, without throwing it. Styles other than whitespace can make ParseExact throw an
    /// <see cref="ArgumentException"/> instead, so they use <paramref name="parse"/>.
    /// </summary>
    public static CellMapperResult InvalidDate(string? text, DateTimeStyles style, string typeName, string[] formats, Action parse)
        => (style & ~DateTimeStyles.AllowWhiteSpaces) != 0
            ? InvalidFromParse(parse)
            : InvalidFormat(text, FormatUtilities.DescribeFormats(typeName, formats), parse);

    /// <summary>
    /// Returns an invalid result for a number that couldn't be parsed, whose exception is created
    /// only if it is needed. It's the exception the type's Parse method throws, created without
    /// throwing it: an <see cref="OverflowException"/> if the text is a number that is out of range
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
    public static CellMapperResult InvalidNumber(string? text, NumberStyles style, IFormatProvider? provider, Type type, bool isInteger, bool canCreateException, Action parse)
        => !canCreateException || text is null || GetNumberFailure(text, style, provider) is not { } isOverflow
            ? InvalidFromParse(parse)
            : CellMapperResult.Invalid(() => isOverflow
                ? new OverflowException($"The value \"{text}\" is too large or too small for {type.Name}{(isInteger ? ", or isn't a whole number" : "")}.")
                : new FormatException($"The value \"{text}\" isn't a valid {type.Name}."));

    // Whether the text a built-in number type's TryParse rejected is a number out of range (true)
    // or not a number (false), or null if that can't be told without parsing.
    private static bool? GetNumberFailure(string text, NumberStyles style, IFormatProvider? provider)
    {
        // Hex and binary numbers aren't parsed by decimal and double.
        if ((style & ~NumberStyles.Any) != 0)
        {
            return null;
        }

        // A number that decimal or double can parse, with the same style, is a number the type
        // can't hold.
        if (decimal.TryParse(text, style, provider, out _))
        {
            return true;
        }

        if (!double.TryParse(text, style, provider, out var number))
        {
            return false;
        }

        // Infinity and NaN symbols, or a number too large for a double.
        return double.IsNaN(number) || double.IsInfinity(number) ? null : true;
    }
}

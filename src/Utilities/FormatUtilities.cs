using System.Linq;
using System.Globalization;
using System.Runtime.CompilerServices;

namespace ExcelMapper.Utilities;

internal static class FormatUtilities
{
    // ISO 8601 dates and times. "K" matches "Z", an offset such as "+01:00", or nothing, and
    // ".FFFFFFF" matches no fraction or up to seven digits.
    private static readonly string[] s_isoDateTimeFormats =
    [
        "yyyy-MM-dd",
        "yyyy-MM-ddTHH:mmK",
        "yyyy-MM-ddTHH:mm:ss.FFFFFFFK",
        "yyyy-MM-dd HH:mmK",
        "yyyy-MM-dd HH:mm:ss.FFFFFFFK",
    ];

    private static readonly ConditionalWeakTable<DateTimeFormatInfo, string[]> s_defaultDateTimeFormats = new();

    /// <summary>
    /// Gets the formats also tried for <see cref="DateTime"/> and <see cref="DateTimeOffset"/> text
    /// when the formats haven't been set: the culture's short date ("d") and short date and time
    /// ("g"), the culture's patterns with ordinary spaces, and ISO 8601.
    /// </summary>
    public static string[] GetDefaultDateTimeFormats(IFormatProvider? provider)
        => GetFormats(provider, s_defaultDateTimeFormats, static info =>
        {
            // ISO 8601 first: it's the most common way to write dates as text, for example in CSV files.
            var formats = new List<string>(s_isoDateTimeFormats) { "g", "d" };
            AddWithOrdinarySpaces(formats, $"{info.ShortDatePattern} {info.LongTimePattern}");
            AddWithOrdinarySpaces(formats, $"{info.ShortDatePattern} {info.ShortTimePattern}");
            return [.. formats];
        });

#if NET6_0_OR_GREATER
    /// <summary>
    /// Formats also tried for <see cref="DateOnly"/> text when the formats haven't been set: ISO 8601.
    /// </summary>
    public static readonly string[] DefaultDateOnlyFormats = ["yyyy-MM-dd"];

    private static readonly ConditionalWeakTable<DateTimeFormatInfo, string[]> s_defaultTimeOnlyFormats = new();

    /// <summary>
    /// Gets the formats also tried for <see cref="TimeOnly"/> text when the formats haven't been set:
    /// the culture's long time ("T"), the culture's patterns with ordinary spaces, and 24-hour times
    /// with optional seconds and fractions.
    /// </summary>
    public static string[] GetDefaultTimeOnlyFormats(IFormatProvider? provider)
        => GetFormats(provider, s_defaultTimeOnlyFormats, static info =>
        {
            var formats = new List<string> { "T" };
            AddWithOrdinarySpaces(formats, info.LongTimePattern);
            AddWithOrdinarySpaces(formats, info.ShortTimePattern);
            formats.Add("H:mm");
            formats.Add("H:mm:ss.FFFFFFF");
            return [.. formats];
        });
#endif

    /// <summary>
    /// Whether text could be a date or time in one of the default formats, which all contain a
    /// digit. Text such as "N/A" is rejected without trying every format.
    /// </summary>
    public static bool HasDigit(string? value)
    {
        if (value is null)
        {
            return false;
        }

        foreach (var c in value)
        {
            if (c is >= '0' and <= '9')
            {
                return true;
            }
        }

        return false;
    }

    private static string[] GetFormats(IFormatProvider? provider, ConditionalWeakTable<DateTimeFormatInfo, string[]> cache, Func<DateTimeFormatInfo, string[]> create)
    {
        var info = DateTimeFormatInfo.GetInstance(provider);

        // Cultures' format info is read-only; a writable one may still change, so don't cache it.
        if (!info.IsReadOnly)
        {
            return create(info);
        }

        return cache.TryGetValue(info, out var formats) ? formats : AddFormats(cache, info, create);
    }

    // A separate method so the closure is only allocated the first time a culture is seen.
    private static string[] AddFormats(ConditionalWeakTable<DateTimeFormatInfo, string[]> cache, DateTimeFormatInfo info, Func<DateTimeFormatInfo, string[]> create)
        => cache.GetValue(info, key => create(key));

    /// <summary>
    /// Adds a culture's pattern with its non-breaking spaces replaced by ordinary spaces. On ICU-based
    /// .NET (Linux, macOS and Windows since .NET 5), cultures such as en-US put U+202F before AM/PM,
    /// so text with an ordinary space, as people and most programs write it, doesn't match.
    /// </summary>
    private static void AddWithOrdinarySpaces(List<string> formats, string pattern)
    {
        var withOrdinarySpaces = pattern.Replace('\u202F', ' ').Replace('\u00A0', ' ');
        if (withOrdinarySpaces != pattern)
        {
            formats.Add(withOrdinarySpaces);
        }
    }

    public static void ValidateFormats(string[] formats, [CallerArgumentExpression(nameof(formats))] string? paramName = null)
    {
        ThrowHelpers.ThrowIfNull(formats, paramName);
        if (formats.Length == 0)
        {
            throw new ArgumentException("Formats cannot be empty.", paramName);
        }
        foreach (var format in formats)
        {
            if (string.IsNullOrEmpty(format))
            {
                throw new ArgumentException("Formats cannot contain null or empty values.", paramName);
            }
        }
    }

    /// <summary>
    /// Describes the formats a value was parsed with, for the exception for an invalid value: for
    /// example 'a valid DateTime in the format "yyyy-MM-dd"'.
    /// </summary>
    public static string DescribeFormats(string typeName, string[] formats)
        => formats.Length == 1
            ? $"a valid {typeName} in the format \"{formats[0]}\""
            : $"a valid {typeName} in one of the formats {string.Join(", ", formats.Select(f => $"\"{f}\""))}";
}

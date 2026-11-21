using System.Globalization;

namespace ExcelMapper;

/// <summary>
/// Specifies the offset from UTC given to cells that Excel stores as dates, when mapping them to a
/// <see cref="DateTimeOffset"/>. Excel dates have no time zone, so by default the machine's local
/// offset is used.
/// </summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
public class ExcelOffsetAttribute : Attribute
{
    /// <summary>
    /// Specifies the offset from UTC, such as "+00:00" (UTC), "-08:00" or "+05:30".
    /// </summary>
    /// <param name="offset">The offset in [+|-]hh:mm form.</param>
    public ExcelOffsetAttribute(string offset)
    {
        ThrowHelpers.ThrowIfNull(offset, nameof(offset));
        var text = offset.Trim();
        var negative = text.StartsWith("-", StringComparison.Ordinal);
        if (text.StartsWith("+", StringComparison.Ordinal) || negative)
        {
            text = text.Substring(1);
        }

        if (!TimeSpan.TryParseExact(text, @"hh\:mm", CultureInfo.InvariantCulture, out var value) ||
            value > TimeSpan.FromHours(14))
        {
            throw new ArgumentException($"Invalid offset '{offset}'. Expected [+|-]hh:mm between -14:00 and +14:00, such as \"+00:00\" or \"-08:00\".", nameof(offset));
        }

        Offset = negative ? value.Negate() : value;
    }

    /// <summary>
    /// Gets the offset from UTC.
    /// </summary>
    public TimeSpan Offset { get; }
}

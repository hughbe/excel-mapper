using System.Globalization;

namespace ExcelMapper.Utilities;

/// <summary>
/// How a pipeline's built-in conversions parse text: the formats, format provider, number style,
/// kind of URI and offset set with the fluent methods (<c>WithFormats</c>, <c>WithFormatProvider</c>,
/// <c>WithNumberStyle</c>, <c>WithUriKind</c>, <c>WithOffset</c>) or their attributes. Each pipeline
/// has its own, and gives it to the mappers with the cell (<see cref="CellContext.Options"/>), so the
/// built-in mappers hold no state and are shared. Null means the mapper's default.
/// </summary>
internal sealed class ConversionOptions
{
    /// <summary>
    /// The options of a pipeline that has none set.
    /// </summary>
    public static ConversionOptions Default { get; } = new();

    private string[]? _formats;
    private UriKind? _uriKind;
    private TimeSpan? _offset;

    /// <summary>
    /// Gets or sets the formats dates and times are parsed with. Null means the type's default
    /// format, then the culture's common formats.
    /// </summary>
    public string[]? Formats
    {
        get => _formats;
        set
        {
            if (value is not null)
            {
                FormatUtilities.ValidateFormats(value);
            }

            _formats = value;
        }
    }

    /// <summary>
    /// Gets or sets the format provider text is parsed with. Null means the sheet's culture.
    /// </summary>
    public IFormatProvider? Provider { get; set; }

    /// <summary>
    /// Gets or sets the number style numbers are parsed with. Null means the type's default.
    /// </summary>
    public NumberStyles? NumberStyle { get; set; }

    /// <summary>
    /// Gets or sets the kind of URI accepted. Null means absolute URIs.
    /// </summary>
    public UriKind? UriKind
    {
        get => _uriKind;
        set
        {
            if (value is { } uriKind)
            {
                EnumUtilities.ValidateIsDefined(uriKind);
            }

            _uriKind = value;
        }
    }

    /// <summary>
    /// Gets or sets the offset of <see cref="DateTimeOffset"/> values read from date cells or from
    /// text without an offset. Null means the local time zone's offset.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The offset is not a whole number of minutes
    /// between -14 and +14 hours.</exception>
    public TimeSpan? Offset
    {
        get => _offset;
        set
        {
            if (value is { } offset && (offset.Ticks % TimeSpan.TicksPerMinute != 0 || offset < TimeSpan.FromHours(-14) || offset > TimeSpan.FromHours(14)))
            {
                throw new ArgumentOutOfRangeException(nameof(value), offset, "The offset must be a whole number of minutes between -14 and +14 hours.");
            }

            _offset = value;
        }
    }
}

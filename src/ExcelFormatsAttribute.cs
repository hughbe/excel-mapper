namespace ExcelMapper;

/// <summary>
/// An attribute used to specify the formats used when mapping date, time or numeric values.
/// </summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
public class ExcelFormatsAttribute : Attribute
{
    /// <summary>
    /// Gets the formats.
    /// </summary>
    public string[] Formats { get; }

    /// <summary>
    /// Constructs the attribute with the specified formats.
    /// </summary>
    /// <param name="formats">The formats.</param>
    public ExcelFormatsAttribute(params string[] formats)
    {
        FormatUtilities.ValidateFormats(formats);
        Formats = formats;
    }
}

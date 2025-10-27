namespace ExcelMapper;

/// <summary>
/// Specifies the separators used when mapping collections from cell values.
/// </summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
public class ExcelSeparatorsAttribute : Attribute
{
    /// <summary>
    /// Gets the string separators.
    /// </summary>
    public string[]? StringSeparators { get; }

    /// <summary>
    /// Gets the char separators.
    /// </summary>
    public char[]? CharSeparators { get; }

    /// <summary>
    /// Gets or sets the string split options.
    /// </summary>
    public StringSplitOptions Options { get; set; } = StringSplitOptions.None;

    /// <summary>
    /// Constructs the attribute with the specified separators.
    /// </summary>
    /// <param name="separators">The separators.</param>
    public ExcelSeparatorsAttribute(params string[] separators)
    {
        SeparatorUtilities.ValidateSeparators(separators);
        StringSeparators = separators;
    }

    /// <summary>
    /// Constructs the attribute with the specified separators.
    /// </summary>
    /// <param name="separators">The separators.</param>
    public ExcelSeparatorsAttribute(params char[] separators)
    {
        SeparatorUtilities.ValidateSeparators(separators);
        CharSeparators = separators;
    }
}

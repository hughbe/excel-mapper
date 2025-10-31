using System.Globalization;

namespace ExcelMapper;

/// <summary>
/// An attribute used to specify the number styles to use when parsing numeric values.
/// </summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
public class ExcelNumberStyleAttribute : Attribute
{
    /// <summary>
    /// Gets the number styles.
    /// </summary>
    public NumberStyles Style { get; }

    /// <summary>
    /// Constructs the attribute with the specified number styles.
    /// </summary>
    /// <param name="style">The number styles to use when parsing numeric values. This determines what formatting elements are allowed (e.g., thousands separators, parentheses for negatives, leading/trailing whitespace).</param>
    public ExcelNumberStyleAttribute(NumberStyles style)
    {
        Style = style;
    }
}

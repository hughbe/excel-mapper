namespace ExcelMapper;

/// <summary>
/// The attribute used to specify the string comparison for dictionary key comparisons.
/// </summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
public class ExcelMappingDictionaryComparerAttribute : Attribute
{
    /// <summary>
    /// The string comparison to use for dictionary key comparisons.
    /// </summary>
    public StringComparison Comparison { get; }

    /// <summary>
    /// Creates a new <see cref="ExcelMappingDictionaryComparerAttribute"/> instance.
    /// </summary>
    /// <param name="comparison">The string comparison to use for dictionary key comparisons.</param>
    /// <exception cref="ArgumentException">Thrown when an invalid comparison is specified.</exception>
    public ExcelMappingDictionaryComparerAttribute(StringComparison comparison)
    {
        EnumUtilities.ValidateIsDefined(comparison);
        Comparison = comparison;
    }
}

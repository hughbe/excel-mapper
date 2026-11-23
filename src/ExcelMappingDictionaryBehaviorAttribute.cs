namespace ExcelMapper;

/// <summary>
/// Specifies the behavior of a Dictionary mapper when mapping keys.
/// </summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
public sealed class ExcelMappingDictionaryBehaviorAttribute : Attribute
{
    /// <summary>
    /// The behavior of the Dictionary mapper when mapping keys.
    /// </summary>
    public MappingDictionaryBehavior Behavior { get; }

    /// <summary>
    /// Creates a new <see cref="ExcelMappingDictionaryBehaviorAttribute"/> instance.
    /// </summary>
    /// <param name="behavior">The behavior of the Dictionary mapper when mapping keys.</param>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when an invalid behavior is specified.</exception>
    public ExcelMappingDictionaryBehaviorAttribute(MappingDictionaryBehavior behavior)
    {
        EnumUtilities.ValidateIsDefined(behavior);
        Behavior = behavior;
    }
}

using ExcelMapper.Mappers;

namespace ExcelMapper;

/// <summary>
/// Specifies the behavior of a Dictionary mapper when mapping keys.
/// </summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
public class ExcelMappingDictionaryBehaviorAttribute : Attribute
{
    /// <summary>
    /// The behavior of the Dictionary mapper when mapping keys.
    /// </summary>
    public MappingDictionaryMapperBehavior Behavior { get; }

    /// <summary>
    /// Creates a new <see cref="ExcelMappingDictionaryBehaviorAttribute"/> instance.
    /// </summary>
    /// <param name="behavior">The behavior of the Dictionary mapper when mapping keys.</param>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when an invalid behavior is specified.</exception>
    public ExcelMappingDictionaryBehaviorAttribute(MappingDictionaryMapperBehavior behavior)
    {
        if (!Enum.IsDefined(behavior))
        {
            throw new ArgumentOutOfRangeException(nameof(behavior), behavior, $"Invalid value \"{behavior}\".");
        }

        Behavior = behavior;
    }
}

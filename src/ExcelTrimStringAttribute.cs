namespace ExcelMapper;

/// <summary>
/// Indicates that string values for the property or field should be trimmed.
/// </summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
public sealed class ExcelTrimStringAttribute : Attribute
{
    /// <summary>
    /// Initializes a new instance of <see cref="ExcelTrimStringAttribute"/>.
    /// </summary>
    public ExcelTrimStringAttribute()
    {
    }
}

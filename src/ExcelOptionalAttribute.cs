namespace ExcelMapper;

/// <summary>
/// Prevents a property from being deserialized.
/// </summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
public sealed class ExcelOptionalAttribute : Attribute
{
    /// <summary>
    /// Initializes a new instance of <see cref="ExcelOptionalAttribute"/>.
    /// </summary>
    public ExcelOptionalAttribute()
    {
    }
}

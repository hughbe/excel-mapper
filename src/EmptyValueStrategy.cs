namespace ExcelMapper;

/// <summary>
/// What an empty cell gives for a member that has no empty fallback of its own: see
/// <see cref="ExcelClassMap.EmptyValueStrategy"/>.
/// </summary>
public enum EmptyValueStrategy
{
    /// <summary>
    /// <c>null</c> for reference types and nullable types, and an error (an
    /// <see cref="ExcelMappingException"/>) for other value types, such as <c>int</c>. The default.
    /// </summary>
    Error,

    /// <summary>
    /// The type's default value: <c>null</c> for reference types and nullable types, and for other
    /// value types their default, such as <c>0</c>.
    /// </summary>
    DefaultValue
}

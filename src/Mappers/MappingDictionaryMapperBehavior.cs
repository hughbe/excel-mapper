namespace ExcelMapper.Mappers;

/// <summary>
/// Defines the behavior of a Dictionary mapper when mapping keys.
/// </summary>
public enum MappingDictionaryMapperBehavior
{
    /// <summary>
    /// A failure to match is not an error.
    /// </summary>
    Optional,

    /// <summary>
    /// A failure to match is an error.
    /// </summary>
    Required
}
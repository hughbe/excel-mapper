namespace ExcelMapper.Abstractions;

/// <summary>
/// Map for dictionary indexer values.
/// </summary>
internal interface IDictionaryIndexerMap : IMap
{
    /// <summary>
    /// Gets the values map.
    /// </summary>
    Dictionary<object, IMap> Values { get; }
}
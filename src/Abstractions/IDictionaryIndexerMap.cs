namespace ExcelMapper.Abstractions;

/// <summary>
/// Map for dictionary indexer values.
/// </summary>
public interface IDictionaryIndexerMap : IMap
{
    /// <summary>
    /// Gets the values map.
    /// </summary>
    Dictionary<object, IMap> Values { get; }
}
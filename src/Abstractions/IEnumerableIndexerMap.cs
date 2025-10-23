namespace ExcelMapper.Abstractions;

/// <summary>
/// Map for enumerable indexer values.
/// </summary>
public interface IEnumerableIndexerMap : IMap
{
    /// <summary>
    /// The list of maps.
    /// </summary>
    Dictionary<int, IMap> Values { get; }
}
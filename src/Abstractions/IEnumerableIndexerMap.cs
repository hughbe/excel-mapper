namespace ExcelMapper.Abstractions;

/// <summary>
/// Map for enumerable indexer values.
/// </summary>
internal interface IEnumerableIndexerMap : IMap
{
    /// <summary>
    /// The list of maps.
    /// </summary>
    Dictionary<int, IMap> Values { get; }
}
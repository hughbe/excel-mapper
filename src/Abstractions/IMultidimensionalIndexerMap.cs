namespace ExcelMapper.Abstractions;

/// <summary>
/// Map for multidimensional indexer values.
/// </summary>
internal interface IMultidimensionalIndexerMap : IMap
{
    /// <summary>
    /// The list of maps.
    /// </summary>
    Dictionary<int[], IMap> Values { get; }
}
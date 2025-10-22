namespace ExcelMapper.Abstractions;

public interface IEnumerableIndexerMap : IMap
{
    /// <summary>
    /// The list of maps.
    /// </summary>
    Dictionary<int, IMap> Values { get; }
}
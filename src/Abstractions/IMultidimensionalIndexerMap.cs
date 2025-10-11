using System.Collections.Generic;

namespace ExcelMapper.Abstractions;

public interface IMultidimensionalIndexerMap : IMap
{
    /// <summary>
    /// The list of maps.
    /// </summary>
    Dictionary<int[], IMap> Values { get; }
}
using System.Collections.Generic;

namespace ExcelMapper.Abstractions;

public interface IDictionaryIndexerMap : IMap
{
    /// <summary>
    /// The list of maps.
    /// </summary>
    Dictionary<object, IMap> Values { get; }
}
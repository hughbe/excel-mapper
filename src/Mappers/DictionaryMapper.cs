using System;
using System.Collections.Generic;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Mappers;

/// <summary>
/// Tries to map the value of a cell to an object using a mapping dictionary.
/// </summary>
public class DictionaryMapper<T> : ICellMapper
{
    private readonly Dictionary<string, T> _mappingDictionary;

    /// <summary>
    /// Gets the dictionary used to map the value of a cell to an object.
    /// </summary>
    public IReadOnlyDictionary<string, T> MappingDictionary => _mappingDictionary;

    /// <summary>
    /// Constructs a mapper that tries to map the value of a cell to an object using a mapping dictionary.
    /// </summary>
    /// <param name="mappingDictionary">The dictionary used to map the value of a cell to an object.</param>
    /// <param name="comparer">The equality comparer used to the value of a cell to an object.</param>
    public DictionaryMapper(IDictionary<string, T> mappingDictionary, IEqualityComparer<string>? comparer)
    {
        if (mappingDictionary == null)
        {
            throw new ArgumentNullException(nameof(mappingDictionary));
        }

        _mappingDictionary = new Dictionary<string, T>(mappingDictionary, comparer);
    }

    public CellMapperResult MapCellValue(ReadCellResult readResult)
    {
        // If we didn't find anything, keep going. This is not necessarily a fatal error.
        if (readResult.StringValue is null || !_mappingDictionary.TryGetValue(readResult.StringValue, out T result))
        {
            return CellMapperResult.Ignore();
        }

        return CellMapperResult.Success(result!);
    }
}

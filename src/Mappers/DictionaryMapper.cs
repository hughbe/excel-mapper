using System;
using System.Collections.Generic;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Mappers;

/// <summary>
/// Tries to map the value of a cell to an object using a mapping dictionary.
/// </summary>
public class DictionaryMapper<T> : ICellMapper
{
    /// <summary>
    /// Gets the dictionary used to map the value of a cell to an object.
    /// </summary>
    public IReadOnlyDictionary<string, T> MappingDictionary { get; }

    /// <summary>
    /// Gets whether or not a failure to match is an error.
    /// </summary>
    public DictionaryMapperBehavior Behavior { get; }

    /// <summary>
    /// Constructs a mapper that tries to map the value of a cell to an object using a mapping dictionary.
    /// </summary>
    /// <param name="mappingDictionary">The dictionary used to map the value of a cell to an object.</param>
    /// <param name="comparer">The equality comparer used to the value of a cell to an object.</param>
    /// <param name="required">Whether or not an error a failure to match is an error.</param>
    public DictionaryMapper(IDictionary<string, T> mappingDictionary, IEqualityComparer<string>? comparer, DictionaryMapperBehavior behavior)
    {
        if (mappingDictionary == null)
        {
            throw new ArgumentNullException(nameof(mappingDictionary));
        }
        if (!Enum.IsDefined(typeof(DictionaryMapperBehavior), behavior))
        {
            throw new ArgumentException($"Invalid value \"{behavior}\".", nameof(behavior));
        }

        MappingDictionary = new Dictionary<string, T>(mappingDictionary, comparer);
        Behavior = behavior;
    }

    public CellMapperResult MapCellValue(ReadCellResult readResult)
    {
        // If we didn't find anything, keep going. This is not necessarily a fatal error.
        var stringValue = readResult.GetString();
        if (stringValue is null || !MappingDictionary.TryGetValue(stringValue, out T result))
        {
            if (Behavior == DictionaryMapperBehavior.Required)
            {
                return CellMapperResult.Invalid(new ExcelMappingException($"No mapping for value \"{stringValue}\"."));
            }

            return CellMapperResult.Ignore();
        }

        return CellMapperResult.Success(result!);
    }
}

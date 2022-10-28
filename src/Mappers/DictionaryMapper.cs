using System.Reflection;
using ExcelMapper.Abstractions;

/// <summary>
/// Tries to map the value of a cell to an object using a mapping dictionary.
/// </summary>
public class DictionaryMapper<T> : ICellValueMapper
{
    /// <summary>
    /// Gets the dictionary used to map the value of a cell to an object.
    /// </summary>
    public IReadOnlyDictionary<string, T> MappingDictionary { get; }

    /// <summary>
    /// Constructs a mapper that tries to map the value of a cell to an object using a mapping dictionary.
    /// </summary>
    /// <param name="mappingDictionary">The dictionary used to map the value of a cell to an object.</param>
    /// <param name="comparer">The equality comparer used to the value of a cell to an object.</param>
    public DictionaryMapper(IDictionary<string, T> mappingDictionary, IEqualityComparer<string> comparer)
    {
        if (mappingDictionary == null)
        {
            throw new ArgumentNullException(nameof(mappingDictionary));
        }

        MappingDictionary = new Dictionary<string, T>(mappingDictionary, comparer);
    }

    public CellValueMapperResult MapCell(ExcelCell cell, CellValueMapperResult previous, MemberInfo member)
    {
        // If we didn't find anything, keep going. This is not necessarily a fatal error.
        string stringValue = previous.Value?.ToString();
        if (stringValue == null || !MappingDictionary.TryGetValue(stringValue, out T result))
        {
            return CellValueMapperResult.Ignore();
        }

        return CellValueMapperResult.Success(result);
    }
}

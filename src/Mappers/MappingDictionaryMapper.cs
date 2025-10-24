namespace ExcelMapper.Mappers;

/// <summary>
/// Tries to map the value of a cell to an object using a mapping dictionary.
/// </summary>
public class MappingDictionaryMapper<T> : ICellMapper
{
    /// <summary>
    /// Gets the dictionary used to map the value of a cell to an object.
    /// </summary>
    public IReadOnlyDictionary<string, T> MappingDictionary { get; }

    /// <summary>
    /// Gets whether or not a failure to match is an error.
    /// </summary>
    public MappingDictionaryMapperBehavior Behavior { get; }

    /// <summary>
    /// Constructs a mapper that tries to map the value of a cell to an object using a mapping dictionary.
    /// </summary>
    /// <param name="mappingDictionary">The dictionary used to map the value of a cell to an object.</param>
    /// <param name="comparer">The equality comparer used to the value of a cell to an object.</param>
    /// <param name="behavior">Whether or not a failure to match is an error.</param>
    public MappingDictionaryMapper(IDictionary<string, T> mappingDictionary, IEqualityComparer<string>? comparer, MappingDictionaryMapperBehavior behavior)
    {
        ArgumentNullException.ThrowIfNull(mappingDictionary);
        if (!Enum.IsDefined(behavior))
        {
            throw new ArgumentException($"Invalid value \"{behavior}\".", nameof(behavior));
        }

        MappingDictionary = new Dictionary<string, T>(mappingDictionary, comparer);
        Behavior = behavior;
    }

    /// <inheritdoc/>
    public CellMapperResult Map(ReadCellResult readResult)
    {
        // If we didn't find anything, keep going. This is not necessarily a fatal error.
        var stringValue = readResult.GetString();
        if (stringValue is null || !MappingDictionary.TryGetValue(stringValue, out T? result))
        {
            if (Behavior == MappingDictionaryMapperBehavior.Required)
            {
                return CellMapperResult.Invalid(new ExcelMappingException($"No mapping for value \"{stringValue}\"."));
            }

            return CellMapperResult.Ignore();
        }

        return CellMapperResult.Success(result!);
    }
}

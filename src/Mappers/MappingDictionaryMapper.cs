namespace ExcelMapper.Mappers;

/// <summary>
/// Maps fixed text to values, for <see cref="ExcelMappingDictionaryAttribute"/> and
/// <c>WithMapping(dictionary)</c>.
/// </summary>
internal sealed class MappingDictionaryMapper<T> : ICellMapper<T>
{
    public IReadOnlyDictionary<string, T> MappingDictionary { get; }

    public MappingDictionaryMapperBehavior Behavior { get; }

    public MappingDictionaryMapper(IDictionary<string, T> mappingDictionary, IEqualityComparer<string>? comparer, MappingDictionaryMapperBehavior behavior)
    {
        ThrowHelpers.ThrowIfNull(mappingDictionary, nameof(mappingDictionary));
        EnumUtilities.ValidateIsDefined(behavior);

        MappingDictionary = new Dictionary<string, T>(mappingDictionary, comparer);
        Behavior = behavior;
    }

    public CellConversion Convert(ReadCellResult cell, out T value)
    {
        var stringValue = cell.GetString();
        if (stringValue is not null && MappingDictionary.TryGetValue(stringValue, out value!))
        {
            return CellConversion.Success;
        }

        // Text that isn't in the dictionary is converted by the other mappers, unless it's required.
        value = default!;
        return Behavior == MappingDictionaryMapperBehavior.Required ? CellConversion.Invalid : CellConversion.NotHandled;
    }

    public Exception CreateException(ReadCellResult cell)
        => new ExcelMappingException($"No mapping for value \"{cell.GetString()}\".");
}

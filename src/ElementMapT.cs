namespace ExcelMapper;

/// <summary>
/// How the elements of a collection, or the values of a dictionary, are converted: configured with
/// <see cref="CollectionMap{TElement}.WithElementMap"/> and <see cref="DictionaryMap{TKey, TValue}.WithValueMap"/>.
/// </summary>
/// <typeparam name="T">The type of the elements or values.</typeparam>
public sealed class ElementMap<T> : ValuePipeline<T, ElementMap<T>>
{
    internal ElementMap()
    {
    }
}

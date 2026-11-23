using ExcelMapper.Mappers;
using ExcelMapper.Transformers;

namespace ExcelMapper;

/// <summary>
/// How a map converts cells to <typeparamref name="T"/>, with value methods that return the map they're called on
/// (<typeparamref name="TSelf"/>), so that calls chain: for example
/// <c>Map(e => e.Name).WithTrim().MakeOptional()</c>. Methods for particular types, such as
/// <c>WithFormats</c>, are in <see cref="ValuePipelineExtensions"/>.
/// </summary>
/// <typeparam name="T">The type of the member, or of the elements or values of a collection.</typeparam>
/// <typeparam name="TSelf">The type of the map: <see cref="ValueMap{T}"/> or <see cref="ElementMap{T}"/>.</typeparam>
public abstract class ValuePipeline<T, TSelf> where TSelf : ValuePipeline<T, TSelf>
{
    private protected ValuePipeline()
    {
    }

    /// <summary>
    /// Gets the pipeline the value methods configure, which converts the cells.
    /// </summary>
    internal ValuePipeline<T> Pipeline { get; } = new();

    private TSelf Self => (TSelf)this;

    /// <summary>
    /// Trims text before it's converted. Cells that Excel stores as numbers or dates keep their type.
    /// </summary>
    /// <returns>The map on which this method was invoked.</returns>
    public TSelf WithTrim()
    {
        Pipeline.Transformers.Add(TrimStringCellTransformer.Instance);
        return Self;
    }

    /// <summary>
    /// Transforms cells with <paramref name="transformers"/>, in order, before they're converted.
    /// The transformers replace the map's existing transformers.
    /// </summary>
    /// <param name="transformers">The transformers.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public TSelf WithTransformers(params ICellTransformer[] transformers)
    {
        ThrowHelpers.ThrowIfNull(transformers, nameof(transformers));
        foreach (var transformer in transformers)
        {
            if (transformer == null)
            {
                throw new ArgumentException("Transformers cannot contain null values.", nameof(transformers));
            }
        }

        Pipeline.Transformers.Clear();
        foreach (var transformer in transformers)
        {
            Pipeline.Transformers.Add(transformer);
        }

        return Self;
    }

    /// <summary>
    /// Maps fixed text to values, for example codes or abbreviations, matched case-sensitively. Text
    /// that isn't in the dictionary is converted as usual.
    /// </summary>
    /// <typeparam name="TValue">The type of the dictionary's values, for example <c>string</c> for a <c>string?</c> member.</typeparam>
    /// <param name="mappingDictionary">The text to map, and the value for each.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public TSelf WithMapping<TValue>(IDictionary<string, TValue> mappingDictionary) where TValue : T
        => WithMapping(mappingDictionary, null, MappingDictionaryBehavior.Optional);

    /// <summary>
    /// Maps fixed text to values, for example codes or abbreviations. Text that isn't in the
    /// dictionary is converted as usual.
    /// </summary>
    /// <typeparam name="TValue">The type of the dictionary's values, for example <c>string</c> for a <c>string?</c> member.</typeparam>
    /// <param name="mappingDictionary">The text to map, and the value for each.</param>
    /// <param name="comparer">How text is compared, for example to ignore case, or null to compare case-sensitively.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public TSelf WithMapping<TValue>(IDictionary<string, TValue> mappingDictionary, IEqualityComparer<string>? comparer) where TValue : T
        => WithMapping(mappingDictionary, comparer, MappingDictionaryBehavior.Optional);

    /// <summary>
    /// Maps fixed text to values, for example codes or abbreviations, matched case-sensitively. Text
    /// that isn't in the dictionary is converted as usual, unless <paramref name="behavior"/> is
    /// <see cref="MappingDictionaryBehavior.Required"/>.
    /// </summary>
    /// <typeparam name="TValue">The type of the dictionary's values, for example <c>string</c> for a <c>string?</c> member.</typeparam>
    /// <param name="mappingDictionary">The text to map, and the value for each.</param>
    /// <param name="behavior">Whether text that isn't in the dictionary is invalid.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public TSelf WithMapping<TValue>(IDictionary<string, TValue> mappingDictionary, MappingDictionaryBehavior behavior) where TValue : T
        => WithMapping(mappingDictionary, null, behavior);

    /// <summary>
    /// Maps fixed text to values, for example codes or abbreviations. Text that isn't in the
    /// dictionary is converted as usual, unless <paramref name="behavior"/> is
    /// <see cref="MappingDictionaryBehavior.Required"/>.
    /// </summary>
    /// <typeparam name="TValue">The type of the dictionary's values, for example <c>string</c> for a <c>string?</c> member.</typeparam>
    /// <param name="mappingDictionary">The text to map, and the value for each.</param>
    /// <param name="comparer">How text is compared, for example to ignore case, or null to compare case-sensitively.</param>
    /// <param name="behavior">Whether text that isn't in the dictionary is invalid.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public TSelf WithMapping<TValue>(IDictionary<string, TValue> mappingDictionary, IEqualityComparer<string>? comparer, MappingDictionaryBehavior behavior) where TValue : T
    {
        ThrowHelpers.ThrowIfNull(mappingDictionary, nameof(mappingDictionary));

        // A dictionary of a type derived from T, such as string for a string? member, is copied.
        if (mappingDictionary is not IDictionary<string, T> values)
        {
            var copy = new Dictionary<string, T>(mappingDictionary.Count);
            foreach (var pair in mappingDictionary)
            {
                copy.Add(pair.Key, pair.Value);
            }

            values = copy;
        }

        Pipeline.Mappers.Add(new MappingDictionaryMapper<T>(values, comparer, behavior));
        return Self;
    }

    /// <summary>
    /// Converts cells with <paramref name="converter"/>, which is given the cell's text. The delegate
    /// replaces the built-in conversion for the member's type and any earlier converter, so it is used
    /// for every cell that isn't empty. Empty cells still use the empty fallback, and values the
    /// delegate throws for use the invalid fallback.
    /// </summary>
    /// <param name="converter">The delegate.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public TSelf WithConverter(Func<string?, T> converter)
    {
        ThrowHelpers.ThrowIfNull(converter, nameof(converter));

        // An ICellConverter<T>, so that values are converted without boxing.
        Pipeline.Mappers.Clear();
        Pipeline.Mappers.Add(new CellConverterMapper<T>(new DelegateCellConverter<T>(converter)));
        return Self;
    }

    /// <summary>
    /// Gives empty cells and invalid cells a fixed value.
    /// </summary>
    /// <param name="value">The value of empty and invalid cells.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public TSelf WithFallback(T? value)
    {
        Pipeline.SetFixedFallback(value, empty: true, invalid: true);
        return Self;
    }

    /// <summary>
    /// Gives empty cells a fixed value.
    /// </summary>
    /// <param name="fallbackValue">The value of empty cells.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public TSelf WithEmptyFallback(T? fallbackValue)
    {
        Pipeline.SetFixedFallback(fallbackValue, empty: true, invalid: false);
        return Self;
    }

    /// <summary>
    /// Gives invalid cells, which can't be converted to the member's type, a fixed value.
    /// </summary>
    /// <param name="fallbackValue">The value of invalid cells.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public TSelf WithInvalidFallback(T? fallbackValue)
    {
        Pipeline.SetFixedFallback(fallbackValue, empty: false, invalid: true);
        return Self;
    }

    /// <summary>
    /// Makes empty cells and invalid cells errors: an <see cref="ExcelMappingException"/> for the cell.
    /// </summary>
    /// <returns>The map on which this method was invoked.</returns>
    public TSelf WithThrowingFallback()
    {
        Pipeline.SetThrowingFallback(empty: true, invalid: true);
        return Self;
    }

    /// <summary>
    /// Makes empty cells errors, for example for a nullable member that mustn't be empty.
    /// </summary>
    /// <returns>The map on which this method was invoked.</returns>
    public TSelf WithThrowingEmptyFallback()
    {
        Pipeline.SetThrowingFallback(empty: true, invalid: false);
        return Self;
    }

    /// <summary>
    /// Makes invalid cells errors, replacing an earlier invalid fallback.
    /// </summary>
    /// <returns>The map on which this method was invoked.</returns>
    public TSelf WithThrowingInvalidFallback()
    {
        Pipeline.SetThrowingFallback(empty: false, invalid: true);
        return Self;
    }
}

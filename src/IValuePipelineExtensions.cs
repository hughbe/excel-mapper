using System.Linq;
using ExcelMapper.Fallbacks;
using ExcelMapper.Mappers;
using ExcelMapper.Transformers;

namespace ExcelMapper;

public delegate TValue ConvertUsingSimpleMapperDelegate<out TValue>(string? stringValue);

/// <summary>
/// Extensions on IValuePipeline to enable fluent "With" method chaining.
/// </summary>
public static class IValuePipelineExtensions
{
    /// <summary>
    /// Specifies that the string value of the cell should be trimmed before it is mapped to
    /// a property or field.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithTrim<TMap>(this TMap map) where TMap : IValuePipeline
    {
        map.Transformers.Add(new TrimStringCellTransformer());
        return map;
    }

    /// <summary>
    /// Specifies additional custom transformers that will be used to map the value of a cell to
    /// a property or field.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="transformers">A list of additional custom transformers that will be used to map the value of a cell to a property or field</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithTransformers<TMap>(this TMap map, params ICellTransformer[] transformers) where TMap : IValuePipeline
        => map.WithTransformers((IEnumerable<ICellTransformer>)transformers);

    /// <summary>
    /// Specifies additional custom transformers that will be used to map the value of a cell to
    /// a property or field.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="transformers">A list of additional custom transformers that will be used to map the value of a cell to a property or field</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithTransformers<TMap>(this TMap map, params IEnumerable<ICellTransformer> transformers) where TMap : IValuePipeline
    {
        ArgumentNullException.ThrowIfNull(transformers);

        foreach (var transformer in transformers)
        {
            if (transformer == null)
            {
                throw new ArgumentException("Transformers cannot contain null values.", nameof(transformers));
            }
        }

        // Clear any existing transformers.
        map.Transformers.Clear();

        // Then, add the new transformers.
        foreach (var transformer in transformers)
        {
            map.Transformers.Add(transformer);
        }

        return map;
    }

    /// <summary>
    /// Specifies additional custom mappers that will be used to map the value of a cell to
    /// a property or field.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="mappers">A list of additional custom mappers that will be used to map the value of a cell to a property or field</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithMappers<TMap>(this TMap map, params ICellMapper[] mappers) where TMap : IValuePipeline
        => map.WithMappers((IEnumerable<ICellMapper>)mappers);
    
    /// <summary>
    /// Specifies additional custom mappers that will be used to map the value of a cell to
    /// a property or field.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="mappers">A list of additional custom mappers that will be used to map the value of a cell to a property or field</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithMappers<TMap>(this TMap map, params IEnumerable<ICellMapper> mappers) where TMap : IValuePipeline
    {
        ArgumentNullException.ThrowIfNull(mappers);

        foreach (var mapper in mappers)
        {
            if (mapper == null)
            {
                throw new ArgumentException("Mappers cannot contain null values.", nameof(mappers));
            }
        }

        // Clear any existing mappers.
        map.Mappers.Clear();

        // Then, add the new mappers.
        foreach (var mapper in mappers)
        {
            map.Mappers.Add(mapper);
        }

        return map;
    }

    /// <summary>
    /// Specifies that the value of a cell should be mapped to a fixed value if it cannot be parsed. This
    /// is useful for mapping columns where equivalent data was entered differently.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <typeparam name="T">The type of the property or field that the map represents.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="mappingDictionary">A dictionary that maps a fixed string value to a fixed value of T.</param>
    /// <param name="comparer">The comparer uses to map fixed string values. This allows for case-insensitive mappings, for example.</param>
    /// <param name="behavior">Whether or not a failure to match is an error.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithMapping<TMap, TValue>(this TMap map, IDictionary<string, TValue> mappingDictionary, IEqualityComparer<string>? comparer = null, MappingDictionaryMapperBehavior behavior = MappingDictionaryMapperBehavior.Optional) where TMap : IValuePipeline<TValue>
    {
        var item = new MappingDictionaryMapper<TValue>(mappingDictionary, comparer, behavior);
        map.Mappers.Add(item);
        return map;
    }

    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a DateTime. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a DateTime.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static IValuePipeline<DateTime> WithFormats(this IValuePipeline<DateTime> map, params string[] formats)
    {
        map.AddFormats<DateTimeMapper>(formats);
        return map;
    }

    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a DateTime. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a DateTime.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static IValuePipeline<DateTime> WithFormats(this IValuePipeline<DateTime> map, IEnumerable<string> formats)
    {
        ArgumentNullException.ThrowIfNull(formats);
        return map.WithFormats([.. formats]);
    }

    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a DateTime. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a DateTime.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static IValuePipeline<DateTime?> WithFormats(this IValuePipeline<DateTime?> map, params string[] formats)
    {
        map.AddFormats<DateTimeMapper>(formats);
        return map;
    }

    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a DateTime. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a DateTime.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static IValuePipeline<DateTime?> WithFormats(this IValuePipeline<DateTime?> map, IEnumerable<string> formats)
    {
        ArgumentNullException.ThrowIfNull(formats);
        return map.WithFormats([.. formats]);
    }

    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a DateTimeOffset. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a DateTimeOffset.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static IValuePipeline<DateTimeOffset> WithFormats(this IValuePipeline<DateTimeOffset> map, params string[] formats)
    {
        map.AddFormats<DateTimeOffsetMapper>(formats);
        return map;
    }

    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a DateTimeOffset. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a DateTimeOffset.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static IValuePipeline<DateTimeOffset> WithFormats(this IValuePipeline<DateTimeOffset> map, IEnumerable<string> formats)
    {
        ArgumentNullException.ThrowIfNull(formats);
        return map.WithFormats([.. formats]);
    }

    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a DateTimeOffset. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a DateTimeOffset.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static IValuePipeline<DateTimeOffset?> WithFormats(this IValuePipeline<DateTimeOffset?> map, params string[] formats)
    {
        map.AddFormats<DateTimeOffsetMapper>(formats);
        return map;
    }

    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a DateTimeOffset. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a DateTimeOffset.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static IValuePipeline<DateTimeOffset?> WithFormats(this IValuePipeline<DateTimeOffset?> map, IEnumerable<string> formats)
    {
        ArgumentNullException.ThrowIfNull(formats);
        return map.WithFormats([.. formats]);
    }

    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a TimeSpan. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a TimeSpan.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static IValuePipeline<TimeSpan> WithFormats(this IValuePipeline<TimeSpan> map, params string[] formats)
    {
        map.AddFormats<TimeSpanMapper>(formats);
        return map;
    }

    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a TimeSpan. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a TimeSpan.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static IValuePipeline<TimeSpan> WithFormats(this IValuePipeline<TimeSpan> map, IEnumerable<string> formats)
    {
        ArgumentNullException.ThrowIfNull(formats);
        return map.WithFormats([.. formats]);
    }

    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a TimeSpan. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a TimeSpan.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static IValuePipeline<TimeSpan?> WithFormats(this IValuePipeline<TimeSpan?> map, params string[] formats)
    {
        map.AddFormats<TimeSpanMapper>(formats);
        return map;
    }

    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a TimeSpan. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a TimeSpan.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static IValuePipeline<TimeSpan?> WithFormats(this IValuePipeline<TimeSpan?> map, IEnumerable<string> formats)
    {
        ArgumentNullException.ThrowIfNull(formats);
        return map.WithFormats([.. formats]);
    }

    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a DateOnly. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a DateOnly.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static IValuePipeline<DateOnly> WithFormats(this IValuePipeline<DateOnly> map, params string[] formats)
    {
        map.AddFormats<DateOnlyMapper>(formats);
        return map;
    }

    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a DateOnly. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a DateOnly.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static IValuePipeline<DateOnly> WithFormats(this IValuePipeline<DateOnly> map, IEnumerable<string> formats)
    {
        ArgumentNullException.ThrowIfNull(formats);
        return map.WithFormats([.. formats]);
    }

    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a DateOnly. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a DateOnly.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static IValuePipeline<DateOnly?> WithFormats(this IValuePipeline<DateOnly?> map, params string[] formats)
    {
        map.AddFormats<DateOnlyMapper>(formats);
        return map;
    }

    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a DateOnly. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a DateOnly.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static IValuePipeline<DateOnly?> WithFormats(this IValuePipeline<DateOnly?> map, IEnumerable<string> formats)
    {
        ArgumentNullException.ThrowIfNull(formats);
        return map.WithFormats([.. formats]);
    }

    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a TimeOnly. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a TimeOnly.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static IValuePipeline<TimeOnly> WithFormats(this IValuePipeline<TimeOnly> map, params string[] formats)
    {
        map.AddFormats<TimeOnlyMapper>(formats);
        return map;
    }

    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a TimeOnly. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a TimeOnly.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static IValuePipeline<TimeOnly> WithFormats(this IValuePipeline<TimeOnly> map, params IEnumerable<string> formats)
    {
        ArgumentNullException.ThrowIfNull(formats);
        return map.WithFormats([.. formats]);
    }

    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a TimeOnly. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a TimeOnly.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static IValuePipeline<TimeOnly?> WithFormats(this IValuePipeline<TimeOnly?> map, params string[] formats)
    {
        map.AddFormats<TimeOnlyMapper>(formats);
        return map;
    }

    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a TimeOnly. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a TimeOnly.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static IValuePipeline<TimeOnly?> WithFormats(this IValuePipeline<TimeOnly?> map, IEnumerable<string> formats)
    {
        ArgumentNullException.ThrowIfNull(formats);
        return map.WithFormats([.. formats]);
    }

    private static void ValidateFormats(string[] formats)
    {
        ArgumentNullException.ThrowIfNull(formats);
        if (formats.Length == 0)
        {
            throw new ArgumentException("Formats cannot be empty.", nameof(formats));
        }
        foreach (var format in formats)
        {
            if (string.IsNullOrEmpty(format))
            {
                throw new ArgumentException("Formats cannot contain null or empty values.", nameof(formats));
            }
        }
    }

    private static void AddFormats<TCellMapper>(this IValuePipeline map, string[] formats) where TCellMapper : IFormatsCellMapper, new()
    {
        ArgumentNullException.ThrowIfNull(formats);
        ValidateFormats(formats);

        var mapper = map.Mappers
            .OfType<TCellMapper>()
            .FirstOrDefault();
        if (mapper == null)
        {
            mapper = new TCellMapper();
            map.Mappers.Add(mapper);
        }

        mapper.Formats = formats;
    }

    /// <summary>
    /// Specifies that the value of a cell should be mapped to a value using the given delegate. This is
    /// useful for specifying custom mapping behaviour for a property or field without having to write
    /// your own ICellMapper.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <typeparam name="TMap">The type of the property or field that the map represents.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="converter">A delegate that is invoked to map the string value of a cell to the value of a property or field.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithConverter<TMap, TValue>(
        this TMap map,
        ConvertUsingSimpleMapperDelegate<TValue?> converter)
        where TMap : IValuePipeline<TValue?>
    {
        ArgumentNullException.ThrowIfNull(converter);

        CellMapperResult actualConverter(ReadCellResult readResult)
        {
            try
            {
                object? result = converter(readResult.GetString());
                return CellMapperResult.Success(result);
            }
            catch (Exception exception)
            {
                return CellMapperResult.Invalid(exception);
            }
        }

        return map.WithConverter(actualConverter);
    }

    /// <summary>
    /// Specifies that the value of a cell should be mapped to a value using the given delegate. This is
    /// useful for specifying custom mapping behaviour for a property or field without having to write
    /// your own ICellMapper.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="converter">A delegate that is invoked to map the read cell result to the value of a property or field.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithConverter<TMap>(this TMap map, ConvertUsingMapperDelegate converter) where TMap: IValuePipeline
    {
        ArgumentNullException.ThrowIfNull(converter);

        var item = new ConvertUsingMapper(converter);
        map.Mappers.Add(item);
        return map;
    }

    /// <summary>
    /// Specifies that the map should throw an exception if the value of a cell if empty or cannot be mapped.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithThrowingFallback<TMap>(this TMap map) where TMap : IValuePipeline
        => map.WithFallbackItem(new ThrowFallback());

    /// <summary>
    /// Specifies a fixed fallback to be used if the value of a cell is empty or cannot be mapped.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="value">The value that will be assigned to the property or field if the value of a cell is empty or cannot be mapped.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithValueFallback<TMap>(this TMap map, object? value) where TMap : IValuePipeline
        => map.WithFallbackItem(new FixedValueFallback(value));

    /// <summary>
    /// Specifies a custom fallback to be used if the value of a cell is empty or cannot be mapped.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="fallbackItem">The fallback to be used if the value of a cell is empty or cannot be mapped.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFallbackItem<TMap>(this TMap map, IFallbackItem fallbackItem) where TMap : IValuePipeline
    {
        ArgumentNullException.ThrowIfNull(fallbackItem);
        map.EmptyFallback = fallbackItem;
        map.InvalidFallback = fallbackItem;
        return map;
    }

    /// <summary>
    /// Specifies a fixed fallback to be used if the value of a cell is empty.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="fallbackValue">The value that will be assigned to the property or field if the value of a cell is empty.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithEmptyFallback<TMap>(this TMap map, object? fallbackValue) where TMap : IValuePipeline
        => map.WithEmptyFallbackItem(new FixedValueFallback(fallbackValue));

    /// <summary>
    /// Specifies a custom fallback to be used if the value of a cell is empty.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="fallbackItem">The fallback to be used if the value of a cell is empty.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithEmptyFallbackItem<TMap>(this TMap map, IFallbackItem fallbackItem) where TMap : IValuePipeline
    {
        ArgumentNullException.ThrowIfNull(fallbackItem);
        map.EmptyFallback = fallbackItem;
        return map;
    }

    /// <summary>
    /// Specifies that the map should throw an exception if the value of a cell is empty.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithThrowingEmptyFallback<TMap>(this TMap map) where TMap : IValuePipeline
        => map.WithEmptyFallbackItem(new ThrowFallback());

    /// <summary>
    /// Specifies that the map should throw an exception if the value of a cell cannot be mapped.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithThrowingInvalidFallback<TMap>(this TMap map) where TMap : IValuePipeline
        => map.WithInvalidFallbackItem(new ThrowFallback());

    /// <summary>
    /// Specifies a fixed fallback to be used if the value of a cell cannot be mapped.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="fallbackValue">The value that will be assigned to the property or field if the value of a cell cannot be mapped.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithInvalidFallback<TMap>(this TMap map, object? fallbackValue) where TMap : IValuePipeline
        => map.WithInvalidFallbackItem(new FixedValueFallback(fallbackValue));

    /// <summary>
    /// Specifies a custom fallback to be used if the value of a cell cannot be mapped.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="fallbackItem">The fallback to be used if the value of a cell cannot be mapped.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithInvalidFallbackItem<TMap>(this TMap map, IFallbackItem fallbackItem) where TMap : IValuePipeline
    {
        ArgumentNullException.ThrowIfNull(fallbackItem);
        map.InvalidFallback = fallbackItem;
        return map;
    }
}

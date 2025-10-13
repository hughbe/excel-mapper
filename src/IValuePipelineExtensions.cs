using System;
using System.Collections.Generic;
using System.Linq;
using ExcelMapper.Abstractions;
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
        map.AddCellValueTransformer(new TrimCellValueTransformer());
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
    public static TMap WithCellValueMappers<TMap>(this TMap map, params ICellMapper[] mappers) where TMap : IValuePipeline
    {
        if (mappers == null)
        {
            throw new ArgumentNullException(nameof(mappers));
        }

        foreach (ICellMapper mapper in mappers)
        {
            if (mapper == null)
            {
                throw new ArgumentNullException(nameof(mappers));
            }
        }

        foreach (ICellMapper mapper in mappers)
        {
            map.AddCellValueMapper(mapper);
        }

        return map;
    }

    /// <summary>
    /// Specifies that the value of a cell should be mapped to a fixed value if it cannot be parsed. This
    /// is useful for mapping columns where equivilent data was entered differently.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <typeparam name="T">The type of the property or field that the map represents.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="mappingDictionary">A dictionary that maps a fixed string value to a fixed value of T.</param>
    /// <param name="comparer">The comparer uses to map fixed string values. This allows for case-insensitive mappings, for example.</param>
    /// <param name="behavior">Whether or not an error a failure to match is an error.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithMapping<TMap, TValue>(this TMap map, IDictionary<string, TValue> mappingDictionary, IEqualityComparer<string>? comparer = null, DictionaryMapperBehavior behavior = DictionaryMapperBehavior.Optional) where TMap : IValuePipeline<TValue>
    {
        var item = new DictionaryMapper<TValue>(mappingDictionary, comparer, behavior);
        map.AddCellValueMapper(item);
        return map;
    }

    /// <summary>
    /// Specifies data formats used when mapping the value of a cell to a DateTime. This is useful for
    /// mapping columns where data formats differ. Existing date formats are overriden.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of date formats to use when mapping the value of a cell to a DateTime.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static IValuePipeline<DateTime> WithDateFormats(this IValuePipeline<DateTime> map, params string[] formats)
    {
        map.AddFormats(formats);
        return map;
    }

    /// <summary>
    /// Specifies data formats used when mapping the value of a cell to a DateTime. This is useful for
    /// mapping columns where data formats differ. Existing date formats are overriden.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of date formats to use when mapping the value of a cell to a DateTime.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static IValuePipeline<DateTime> WithDateFormats(this IValuePipeline<DateTime> map, IEnumerable<string> formats)
    {
        if (formats == null)
        {
            throw new ArgumentNullException(nameof(formats));
        }
        
        return map.WithDateFormats([.. formats]);
    }

    /// <summary>
    /// Specifies data formats used when mapping the value of a cell to a DateTime. This is useful for
    /// mapping columns where data formats differ. Existing date formats are overriden.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of date formats to use when mapping the value of a cell to a DateTime.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static IValuePipeline<DateTime?> WithDateFormats(this IValuePipeline<DateTime?> map, params string[] formats)
    {
        map.AddFormats(formats);
        return map;
    }

    /// <summary>
    /// Specifies data formats used when mapping the value of a cell to a DateTime. This is useful for
    /// mapping columns where data formats differ. Existing date formats are overriden.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of date formats to use when mapping the value of a cell to a DateTime.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static IValuePipeline<DateTime?> WithDateFormats(this IValuePipeline<DateTime?> map, IEnumerable<string> formats)
    {
        if (formats == null)
        {
            throw new ArgumentNullException(nameof(formats));
        }
        
        return map.WithDateFormats([.. formats]);
    }

    private static void AddFormats(this IValuePipeline map, string[] formats)
    {
        if (formats == null)
        {
            throw new ArgumentNullException(nameof(formats));
        }
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

        var dateTimeItem = map.CellValueMappers
            .OfType<DateTimeMapper>()
            .FirstOrDefault();
        if (dateTimeItem == null)
        {
            dateTimeItem = new DateTimeMapper();
            map.AddCellValueMapper(dateTimeItem);
        }

        dateTimeItem.Formats = formats;
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
        if (converter == null)
        {
            throw new ArgumentNullException(nameof(converter));
        }

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

    public static TMap WithConverter<TMap>(this TMap map, ConvertUsingMapperDelegate converter) where TMap: IValuePipeline
    {
        if (converter == null)
        {
            throw new ArgumentNullException(nameof(converter));
        }

        var item = new ConvertUsingMapper(converter);
        map.AddCellValueMapper(item);
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
        map.EmptyFallback = fallbackItem ?? throw new ArgumentNullException(nameof(fallbackItem));
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
        map.EmptyFallback = fallbackItem ?? throw new ArgumentNullException(nameof(fallbackItem));
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
        map.InvalidFallback = fallbackItem ?? throw new ArgumentNullException(nameof(fallbackItem));
        return map;
    }
}

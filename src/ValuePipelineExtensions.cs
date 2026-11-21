using System.Globalization;
using System.Linq;
using System.Numerics;
using System.Runtime.InteropServices;
using ExcelMapper.Fallbacks;
using ExcelMapper.Mappers;
using ExcelMapper.Transformers;

namespace ExcelMapper;

/// <summary>
/// A delegate that converts a string value read from an Excel cell to a strongly-typed value.
/// </summary>
/// <typeparam name="TValue">The type of value to convert to.</typeparam>
/// <param name="stringValue">The string value read from the Excel cell, or null if the cell is empty.</param>
/// <returns>The converted value of type <typeparamref name="TValue"/>.</returns>
public delegate TValue ConvertUsingSimpleMapperDelegate<out TValue>(string? stringValue);

/// <summary>
/// Fluent methods for value pipelines: a member's map (<see cref="OneToOneMap{T}"/>) and the
/// element or value pipelines of collections and dictionaries. Each returns the pipeline it was
/// called on.
/// </summary>
public static class ValuePipelineExtensions
{
    /// <summary>
    /// Specifies that the string value of the cell should be trimmed before it is mapped to
    /// a property or field.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithTrim<TMap>(this TMap map) where TMap : ValuePipeline
    {
        map.Transformers.Add(new TrimStringCellTransformer());
        return map;
    }

    /// <summary>
    /// Transforms cells with <paramref name="transformers"/>, in order, before they're converted.
    /// The transformers replace the map's existing transformers.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="transformers">The transformers.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithTransformers<TMap>(this TMap map, params ICellTransformer[] transformers) where TMap : ValuePipeline
    {
        ThrowHelpers.ThrowIfNull(transformers, nameof(transformers));
        foreach (var transformer in transformers)
        {
            if (transformer == null)
            {
                throw new ArgumentException("Transformers cannot contain null values.", nameof(transformers));
            }
        }

        map.Transformers.Clear();
        foreach (var transformer in transformers)
        {
            map.Transformers.Add(new CellTransformerAdapter(transformer));
        }

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
    public static TMap WithTransformers<TMap>(this TMap map, params Abstractions.ICellTransformer[] transformers) where TMap : ValuePipeline
        => map.WithTransformers((IEnumerable<Abstractions.ICellTransformer>)transformers);

    /// <summary>
    /// Specifies additional custom transformers that will be used to map the value of a cell to
    /// a property or field.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="transformers">A list of additional custom transformers that will be used to map the value of a cell to a property or field</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithTransformers<TMap>(this TMap map, params IEnumerable<Abstractions.ICellTransformer> transformers) where TMap : ValuePipeline
    {
        ThrowHelpers.ThrowIfNull(transformers, nameof(transformers));

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
    public static TMap WithMappers<TMap>(this TMap map, params ICellMapper[] mappers) where TMap : ValuePipeline
        => map.WithMappers((IEnumerable<ICellMapper>)mappers);
    
    /// <summary>
    /// Specifies additional custom mappers that will be used to map the value of a cell to
    /// a property or field.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="mappers">A list of additional custom mappers that will be used to map the value of a cell to a property or field</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithMappers<TMap>(this TMap map, params IEnumerable<ICellMapper> mappers) where TMap : ValuePipeline
    {
        ThrowHelpers.ThrowIfNull(mappers, nameof(mappers));

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
    public static TMap WithMapping<TMap, TValue>(this TMap map, IDictionary<string, TValue> mappingDictionary, IEqualityComparer<string>? comparer = null, MappingDictionaryMapperBehavior behavior = MappingDictionaryMapperBehavior.Optional) where TMap : ValuePipeline<TValue>
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
    public static ValuePipeline<DateTime> WithFormats(this ValuePipeline<DateTime> map, params string[] formats)
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
    public static ValuePipeline<DateTime> WithFormats(this ValuePipeline<DateTime> map, params IEnumerable<string> formats)
    {
        ThrowHelpers.ThrowIfNull(formats, nameof(formats));
        return map.WithFormats([.. formats]);
    }

    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a DateTime. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a DateTime.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static ValuePipeline<DateTime?> WithFormats(this ValuePipeline<DateTime?> map, params string[] formats)
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
    public static ValuePipeline<DateTime?> WithFormats(this ValuePipeline<DateTime?> map, params IEnumerable<string> formats)
    {
        ThrowHelpers.ThrowIfNull(formats, nameof(formats));
        return map.WithFormats([.. formats]);
    }

    /// <summary>
    /// Specifies the offset from UTC given to cells that Excel stores as dates. Excel dates have no
    /// time zone, so by default the machine's local offset is used. Use <see cref="TimeSpan.Zero"/> for UTC.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="offset">The offset, a whole number of minutes between -14 and +14 hours.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static ValuePipeline<DateTimeOffset> WithOffset(this ValuePipeline<DateTimeOffset> map, TimeSpan offset)
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.GetOrAddMapper<DateTimeOffsetMapper>().Offset = offset;
        return map;
    }

    /// <summary>
    /// Specifies the offset from UTC given to cells that Excel stores as dates. Excel dates have no
    /// time zone, so by default the machine's local offset is used. Use <see cref="TimeSpan.Zero"/> for UTC.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="offset">The offset, a whole number of minutes between -14 and +14 hours.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static ValuePipeline<DateTimeOffset?> WithOffset(this ValuePipeline<DateTimeOffset?> map, TimeSpan offset)
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.GetOrAddMapper<DateTimeOffsetMapper>().Offset = offset;
        return map;
    }

    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a DateTimeOffset. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a DateTimeOffset.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static ValuePipeline<DateTimeOffset> WithFormats(this ValuePipeline<DateTimeOffset> map, params string[] formats)
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
    public static ValuePipeline<DateTimeOffset> WithFormats(this ValuePipeline<DateTimeOffset> map, params IEnumerable<string> formats)
    {
        ThrowHelpers.ThrowIfNull(formats, nameof(formats));
        return map.WithFormats([.. formats]);
    }

    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a DateTimeOffset. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a DateTimeOffset.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static ValuePipeline<DateTimeOffset?> WithFormats(this ValuePipeline<DateTimeOffset?> map, params string[] formats)
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
    public static ValuePipeline<DateTimeOffset?> WithFormats(this ValuePipeline<DateTimeOffset?> map, params IEnumerable<string> formats)
    {
        ThrowHelpers.ThrowIfNull(formats, nameof(formats));
        return map.WithFormats([.. formats]);
    }

    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a TimeSpan. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a TimeSpan.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static ValuePipeline<TimeSpan> WithFormats(this ValuePipeline<TimeSpan> map, params string[] formats)
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
    public static ValuePipeline<TimeSpan> WithFormats(this ValuePipeline<TimeSpan> map, params IEnumerable<string> formats)
    {
        ThrowHelpers.ThrowIfNull(formats, nameof(formats));
        return map.WithFormats([.. formats]);
    }

    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a TimeSpan. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a TimeSpan.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static ValuePipeline<TimeSpan?> WithFormats(this ValuePipeline<TimeSpan?> map, params string[] formats)
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
    public static ValuePipeline<TimeSpan?> WithFormats(this ValuePipeline<TimeSpan?> map, params IEnumerable<string> formats)
    {
        ThrowHelpers.ThrowIfNull(formats, nameof(formats));
        return map.WithFormats([.. formats]);
    }

#if NET6_0_OR_GREATER
    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a DateOnly. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a DateOnly.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static ValuePipeline<DateOnly> WithFormats(this ValuePipeline<DateOnly> map, params string[] formats)
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
    public static ValuePipeline<DateOnly> WithFormats(this ValuePipeline<DateOnly> map, params IEnumerable<string> formats)
    {
        ThrowHelpers.ThrowIfNull(formats, nameof(formats));
        return map.WithFormats([.. formats]);
    }

    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a DateOnly. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a DateOnly.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static ValuePipeline<DateOnly?> WithFormats(this ValuePipeline<DateOnly?> map, params string[] formats)
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
    public static ValuePipeline<DateOnly?> WithFormats(this ValuePipeline<DateOnly?> map, params IEnumerable<string> formats)
    {
        ThrowHelpers.ThrowIfNull(formats, nameof(formats));
        return map.WithFormats([.. formats]);
    }

    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a TimeOnly. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a TimeOnly.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static ValuePipeline<TimeOnly> WithFormats(this ValuePipeline<TimeOnly> map, params string[] formats)
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
    public static ValuePipeline<TimeOnly> WithFormats(this ValuePipeline<TimeOnly> map, params IEnumerable<string> formats)
    {
        ThrowHelpers.ThrowIfNull(formats, nameof(formats));
        return map.WithFormats([.. formats]);
    }

    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a TimeOnly. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a TimeOnly.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static ValuePipeline<TimeOnly?> WithFormats(this ValuePipeline<TimeOnly?> map, params string[] formats)
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
    public static ValuePipeline<TimeOnly?> WithFormats(this ValuePipeline<TimeOnly?> map, params IEnumerable<string> formats)
    {
        ThrowHelpers.ThrowIfNull(formats, nameof(formats));
        return map.WithFormats([.. formats]);
    }
#endif

    private static void AddFormats<TCellMapper>(this IValuePipeline map, string[] formats) where TCellMapper : IFormatsCellMapper, new()
    {
        FormatUtilities.ValidateFormats(formats);
        var mapper = map.GetOrAddMapper<TCellMapper>();
        mapper.Formats = formats;
    }

    /// <summary>
    /// Specifies the format provider used when mapping the value of a cell to a DateTime. This is useful for
    /// mapping columns where cultures differ. Existing provider is overridden.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider to use when mapping the value of a cell to a DateTime.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static ValuePipeline<DateTime> WithFormatProvider(this ValuePipeline<DateTime> map, IFormatProvider? provider)
    {
        map.AddFormatProvider<DateTimeMapper>(provider);
        return map;
    }

    /// <summary>
    /// Specifies the format provider used when mapping the value of a cell to a DateTime. This is useful for
    /// mapping columns where cultures differ. Existing provider is overridden.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider to use when mapping the value of a cell to a DateTime.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static ValuePipeline<DateTime?> WithFormatProvider(this ValuePipeline<DateTime?> map, IFormatProvider? provider)
    {
        map.AddFormatProvider<DateTimeMapper>(provider);
        return map;
    }

    /// <summary>
    /// Specifies the format provider used when mapping the value of a cell to a DateTimeOffset. This is useful for
    /// mapping columns where cultures differ. Existing provider is overridden.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider to use when mapping the value of a cell to a DateTimeOffset.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static ValuePipeline<DateTimeOffset> WithFormatProvider(this ValuePipeline<DateTimeOffset> map, IFormatProvider? provider)
    {
        map.AddFormatProvider<DateTimeOffsetMapper>(provider);
        return map;
    }

    /// <summary>
    /// Specifies the format provider used when mapping the value of a cell to a DateTimeOffset. This is useful for
    /// mapping columns where cultures differ. Existing provider is overridden.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider to use when mapping the value of a cell to a DateTimeOffset.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static ValuePipeline<DateTimeOffset?> WithFormatProvider(this ValuePipeline<DateTimeOffset?> map, IFormatProvider? provider)
    {
        map.AddFormatProvider<DateTimeOffsetMapper>(provider);
        return map;
    }

    /// <summary>
    /// Specifies the format provider used when mapping the value of a cell to a TimeSpan. This is useful for
    /// mapping columns where cultures differ. Existing provider is overridden.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider to use when mapping the value of a cell to a TimeSpan.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static ValuePipeline<TimeSpan> WithFormatProvider(this ValuePipeline<TimeSpan> map, IFormatProvider? provider)
    {
        map.AddFormatProvider<TimeSpanMapper>(provider);
        return map;
    }

    /// <summary>
    /// Specifies the format provider used when mapping the value of a cell to a TimeSpan. This is useful for
    /// mapping columns where cultures differ. Existing provider is overridden.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider to use when mapping the value of a cell to a TimeSpan.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static ValuePipeline<TimeSpan?> WithFormatProvider(this ValuePipeline<TimeSpan?> map, IFormatProvider? provider)
    {
        map.AddFormatProvider<TimeSpanMapper>(provider);
        return map;
    }

#if NET6_0_OR_GREATER
    /// <summary>
    /// Specifies the format provider used when mapping the value of a cell to a DateOnly. This is useful for
    /// mapping columns where cultures differ. Existing provider is overridden.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider to use when mapping the value of a cell to a DateOnly.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static ValuePipeline<DateOnly> WithFormatProvider(this ValuePipeline<DateOnly> map, IFormatProvider? provider)
    {
        map.AddFormatProvider<DateOnlyMapper>(provider);
        return map;
    }

    /// <summary>
    /// Specifies the format provider used when mapping the value of a cell to a DateOnly. This is useful for
    /// mapping columns where cultures differ. Existing provider is overridden.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider to use when mapping the value of a cell to a DateOnly.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static ValuePipeline<DateOnly?> WithFormatProvider(this ValuePipeline<DateOnly?> map, IFormatProvider? provider)
    {
        map.AddFormatProvider<DateOnlyMapper>(provider);
        return map;
    }

    /// <summary>
    /// Specifies the format provider used when mapping the value of a cell to a TimeOnly. This is useful for
    /// mapping columns where cultures differ. Existing provider is overridden.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider to use when mapping the value of a cell to a TimeOnly.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static ValuePipeline<TimeOnly> WithFormatProvider(this ValuePipeline<TimeOnly> map, IFormatProvider? provider)
    {
        map.AddFormatProvider<TimeOnlyMapper>(provider);
        return map;
    }

    /// <summary>
    /// Specifies the format provider used when mapping the value of a cell to a TimeOnly. This is useful for
    /// mapping columns where cultures differ. Existing provider is overridden.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider to use when mapping the value of a cell to a TimeOnly.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static ValuePipeline<TimeOnly?> WithFormatProvider(this ValuePipeline<TimeOnly?> map, IFormatProvider? provider)
    {
        map.AddFormatProvider<TimeOnlyMapper>(provider);
        return map;
    }
#endif

    private static void AddFormatProvider<TCellMapper>(this IValuePipeline map, IFormatProvider? provider) where TCellMapper : IFormatProviderCellMapper, new()
    {
        var mapper = map.GetOrAddMapper<TCellMapper>();
        mapper.Provider = provider;
    }

#if NET7_0_OR_GREATER
    /// <summary>
    /// Specifies the format provider used when mapping the value of a cell to a numeric type
    /// that implements <see cref="INumberBase{T}"/>. This is useful for mapping columns where
    /// cultures differ. Existing provider is overridden.
    /// </summary>
    /// <typeparam name="T">The type of the property or field that the map represents.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider to use when mapping the value of a cell to a numeric type.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static ValuePipeline<T> WithFormatProvider<T>(this ValuePipeline<T> map, IFormatProvider? provider) where T : INumberBase<T>
    {
        var mapper = map.GetOrAddMapper<INumberBaseMapper<T>>();
        mapper.Provider = provider;
        return map;
    }

    /// <summary>
    /// Specifies the format provider used when mapping the value of a cell to a numeric type
    /// that implements <see cref="INumberBase{T}"/>. This is useful for mapping columns where
    /// cultures differ. Existing provider is overridden.
    /// </summary>
    /// <typeparam name="T">The type of the property or field that the map represents.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider to use when mapping the value of a cell to a numeric type.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static ValuePipeline<T?> WithFormatProvider<T>(this ValuePipeline<T?> map, IFormatProvider? provider) where T : struct, INumberBase<T>
    {
        var mapper = map.GetOrAddMapper<INumberBaseMapper<T>>();
        mapper.Provider = provider;
        return map;
    }

    /// <summary>
    /// Specifies the number styles to use when mapping the value of a cell to a numeric type
    /// that implements <see cref="INumberBase{T}"/>. This is useful for specifying
    /// how numeric values are formatted in the Excel file.
    /// </summary>
    /// <typeparam name="T">The type of the property or field that the map represents.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="style">The number styles to use when mapping the value of a cell to a numeric type.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static ValuePipeline<T> WithNumberStyle<T>(this ValuePipeline<T> map, NumberStyles style) where T : INumberBase<T>
    {
        var mapper = map.GetOrAddMapper<INumberBaseMapper<T>>();
        mapper.Style = style;
        return map;
    }

    /// <summary>
    /// Specifies the number styles to use when mapping the value of a cell to a numeric type
    /// that implements <see cref="INumberBase{T}"/>. This is useful for specifying
    /// how numeric values are formatted in the Excel file.
    /// </summary>
    /// <typeparam name="T">The type of the property or field that the map represents.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="style">The number styles to use when mapping the value of a cell to a numeric type.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static ValuePipeline<T?> WithNumberStyle<T>(this ValuePipeline<T?> map, NumberStyles style) where T : struct, INumberBase<T>
    {
        var mapper = map.GetOrAddMapper<INumberBaseMapper<T>>();
        mapper.Style = style;
        return map;
    }
#endif

    /// <summary>
    /// Specifies the kind of URI to use when mapping the value of a cell to a Uri. This is useful for
    /// specifying whether URIs should be absolute, relative, or either.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="uriKind">The kind of URI to use when mapping the value of a cell to a Uri. Must be a valid UriKind value (Absolute, Relative, or RelativeOrAbsolute).</param>
    /// <returns>The map on which this method was invoked.</returns>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="uriKind"/> is not a valid UriKind value.</exception>
    public static TMap WithUriKind<TMap>(this TMap map, UriKind uriKind) where TMap : ValuePipeline<Uri?>
    {
        EnumUtilities.ValidateIsDefined(uriKind);
        var mapper = map.GetOrAddMapper<UriMapper>();
        mapper.UriKind = uriKind;
        return map;
    }

    private static TCellMapper GetOrAddMapper<TCellMapper>(this IValuePipeline map) where TCellMapper : ICellMapper, new()
    {
        var mapper = map.Mappers
            .OfType<TCellMapper>()
            .FirstOrDefault();
        if (mapper == null)
        {
            mapper = new TCellMapper();
            map.Mappers.Add(mapper);
        }

        return mapper;
    }

    /// <summary>
    /// Specifies that the value of a cell should be mapped to a value using the given delegate. This is
    /// useful for specifying custom mapping behaviour for a property or field without having to write
    /// your own ICellMapper. The delegate replaces the map's mappers, including the built-in conversion
    /// for the member's type and any earlier converter, so it is used for every cell that isn't empty.
    /// Empty cells still use the empty fallback, and values the delegate throws for use the invalid
    /// fallback.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <typeparam name="TMap">The type of the property or field that the map represents.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="converter">A delegate that is invoked to map the string value of a cell to the value of a property or field.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithConverter<TMap, TValue>(
        this TMap map,
        ConvertUsingSimpleMapperDelegate<TValue?> converter)
        where TMap : ValuePipeline<TValue?>
    {
        ThrowHelpers.ThrowIfNull(converter, nameof(converter));

        // An ICellConverter<T>, so that values are converted without boxing.
        map.Mappers.Clear();
        map.Mappers.Add(new CellConverterMapper<TValue?>(new DelegateCellConverter<TValue?>(converter)));
        return map;
    }

    /// <summary>
    /// Specifies that the value of a cell should be mapped to a value using the given delegate. This is
    /// useful for specifying custom mapping behaviour for a property or field without having to write
    /// your own ICellMapper. The delegate replaces the map's mappers, including the built-in conversion
    /// for the member's type and any earlier converter, so it is used for every cell that isn't empty.
    /// Empty cells still use the empty fallback, and values the delegate throws for use the invalid
    /// fallback.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="converter">A delegate that is invoked to map the read cell result to the value of a property or field.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithConverter<TMap>(this TMap map, ConvertUsingMapperDelegate converter) where TMap : ValuePipeline
    {
        ThrowHelpers.ThrowIfNull(converter, nameof(converter));

        // Replace the built-in conversion, rather than adding to it: otherwise the converter would only
        // be used for values the built-in conversion can't convert (for example for numbers), but for
        // every value of other types (such as strings).
        map.Mappers.Clear();
        map.Mappers.Add(new ConvertUsingMapper(converter));
        return map;
    }

    /// <summary>
    /// Converts cells with <paramref name="converter"/>. The converter replaces the map's mappers,
    /// including the built-in conversion for the member's type and any earlier converter, so it is
    /// used for every cell that isn't empty. Empty cells still use the empty fallback, and values the
    /// converter returns <see cref="CellConversionResult{T}.Invalid()"/> for use the invalid fallback.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <typeparam name="TValue">The type of the property or field that the map represents.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="converter">The converter.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithConverter<TMap, TValue>(this TMap map, ICellConverter<TValue> converter) where TMap : ValuePipeline<TValue>
    {
        ThrowHelpers.ThrowIfNull(converter, nameof(converter));

        map.Mappers.Clear();
        map.Mappers.Add(new CellConverterMapper<TValue>(converter));
        return map;
    }

    /// <summary>
    /// Converts the cells of a nullable member with a converter for its underlying type. The converter
    /// replaces the map's mappers, as for <see cref="WithConverter{TMap, TValue}(TMap, ICellConverter{TValue})"/>.
    /// Empty cells still use the empty fallback, which is null by default.
    /// </summary>
    /// <typeparam name="TValue">The underlying type of the property or field that the map represents.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="converter">The converter.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static ValuePipeline<TValue?> WithConverter<TValue>(this ValuePipeline<TValue?> map, ICellConverter<TValue> converter) where TValue : struct
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        ThrowHelpers.ThrowIfNull(converter, nameof(converter));

        map.Mappers.Clear();
        map.Mappers.Add(new CellConverterMapper<TValue?>(new NullableCellConverter<TValue>(converter)));
        return map;
    }

    /// <summary>
    /// Specifies that the map should throw an exception if the value of a cell if empty or cannot be mapped.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithThrowingFallback<TMap>(this TMap map) where TMap : ValuePipeline
        => map.WithFallbackItem(new ThrowFallback());

    /// <summary>
    /// Specifies a fixed fallback to be used if the value of a cell is empty or cannot be mapped.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="value">The value that will be assigned to the property or field if the value of a cell is empty or cannot be mapped.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithValueFallback<TMap>(this TMap map, object? value) where TMap : ValuePipeline
        => map.WithFallbackItem(new FixedValueFallback(value));

    /// <summary>
    /// Specifies a custom fallback to be used if the value of a cell is empty or cannot be mapped.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="fallbackItem">The fallback to be used if the value of a cell is empty or cannot be mapped.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFallbackItem<TMap>(this TMap map, IFallbackItem fallbackItem) where TMap : ValuePipeline
    {
        ThrowHelpers.ThrowIfNull(fallbackItem, nameof(fallbackItem));
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
    public static TMap WithEmptyFallback<TMap>(this TMap map, object? fallbackValue) where TMap : ValuePipeline
        => map.WithEmptyFallbackItem(new FixedValueFallback(fallbackValue));

    /// <summary>
    /// Gives empty cells the value <paramref name="fallback"/> returns.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <typeparam name="TValue">The type of the property or field that the map represents.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="fallback">The fallback.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithEmptyFallback<TMap, TValue>(this TMap map, ICellFallback<TValue> fallback) where TMap : ValuePipeline<TValue>
    {
        ThrowHelpers.ThrowIfNull(fallback, nameof(fallback));
        map.EmptyFallback = new CellFallbackItem<TValue>(fallback);
        return map;
    }

    /// <summary>
    /// Gives invalid cells the value <paramref name="fallback"/> returns.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <typeparam name="TValue">The type of the property or field that the map represents.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="fallback">The fallback.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithInvalidFallback<TMap, TValue>(this TMap map, ICellFallback<TValue> fallback) where TMap : ValuePipeline<TValue>
    {
        ThrowHelpers.ThrowIfNull(fallback, nameof(fallback));
        map.InvalidFallback = new CellFallbackItem<TValue>(fallback);
        return map;
    }

    /// <summary>
    /// Specifies a custom fallback to be used if the value of a cell is empty.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="fallbackItem">The fallback to be used if the value of a cell is empty.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithEmptyFallbackItem<TMap>(this TMap map, IFallbackItem fallbackItem) where TMap : ValuePipeline
    {
        ThrowHelpers.ThrowIfNull(fallbackItem, nameof(fallbackItem));
        map.EmptyFallback = fallbackItem;
        return map;
    }

    /// <summary>
    /// Specifies that the map should throw an exception if the value of a cell is empty.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithThrowingEmptyFallback<TMap>(this TMap map) where TMap : ValuePipeline
        => map.WithEmptyFallbackItem(new ThrowFallback());

    /// <summary>
    /// Specifies that the map should throw an exception if the value of a cell cannot be mapped.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithThrowingInvalidFallback<TMap>(this TMap map) where TMap : ValuePipeline
        => map.WithInvalidFallbackItem(new ThrowFallback());

    /// <summary>
    /// Specifies a fixed fallback to be used if the value of a cell cannot be mapped.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="fallbackValue">The value that will be assigned to the property or field if the value of a cell cannot be mapped.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithInvalidFallback<TMap>(this TMap map, object? fallbackValue) where TMap : ValuePipeline
        => map.WithInvalidFallbackItem(new FixedValueFallback(fallbackValue));

    /// <summary>
    /// Specifies a custom fallback to be used if the value of a cell cannot be mapped.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="fallbackItem">The fallback to be used if the value of a cell cannot be mapped.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithInvalidFallbackItem<TMap>(this TMap map, IFallbackItem fallbackItem) where TMap : ValuePipeline
    {
        ThrowHelpers.ThrowIfNull(fallbackItem, nameof(fallbackItem));
        map.InvalidFallback = fallbackItem;
        return map;
    }
}

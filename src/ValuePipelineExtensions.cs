#if NET7_0_OR_GREATER
using System.Globalization;
using System.Numerics;
#endif
using ExcelMapper.Fallbacks;
using ExcelMapper.Mappers;

namespace ExcelMapper;

/// <summary>
/// The value methods for particular types, such as <c>WithFormats</c> for dates, and the methods
/// that take an <see cref="ICellConverter{T}"/> or <see cref="ICellFallback{T}"/>. They work on a
/// member's map (<see cref="ValueMap{T}"/>) and on the elements or values of a collection
/// (<see cref="ElementMap{T}"/>), and each returns the map it was called on. The methods for any
/// type are on <see cref="ValuePipeline{T, TSelf}"/>.
/// </summary>
public static partial class ValuePipelineExtensions
{
    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a DateTime. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a DateTime.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormats<TMap>(this ValuePipeline<DateTime, TMap> map, params string[] formats) where TMap : ValuePipeline<DateTime, TMap>
    {
        map.Pipeline.SetFormats(formats);
        map.Pipeline.AddMapperIfMissing(DateTimeMapper.Instance);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a DateTime. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a DateTime.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormats<TMap>(this ValuePipeline<DateTime, TMap> map, IEnumerable<string> formats) where TMap : ValuePipeline<DateTime, TMap>
    {
        ThrowHelpers.ThrowIfNull(formats, nameof(formats));
        return map.WithFormats([.. formats]);
    }

    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a DateTime. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a DateTime.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormats<TMap>(this ValuePipeline<DateTime?, TMap> map, params string[] formats) where TMap : ValuePipeline<DateTime?, TMap>
    {
        map.Pipeline.SetFormats(formats);
        map.Pipeline.AddMapperIfMissing(DateTimeMapper.Instance);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a DateTime. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a DateTime.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormats<TMap>(this ValuePipeline<DateTime?, TMap> map, IEnumerable<string> formats) where TMap : ValuePipeline<DateTime?, TMap>
    {
        ThrowHelpers.ThrowIfNull(formats, nameof(formats));
        return map.WithFormats([.. formats]);
    }

    /// <summary>
    /// Specifies the offset from UTC given to dates without one: cells that Excel stores as dates
    /// (which have no time zone), and text without an offset. By default they're UTC. Text with
    /// its own offset keeps it.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="offset">The offset, a whole number of minutes between -14 and +14 hours.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithOffset<TMap>(this ValuePipeline<DateTimeOffset, TMap> map, TimeSpan offset) where TMap : ValuePipeline<DateTimeOffset, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.Offset = offset;
        map.Pipeline.AddMapperIfMissing(DateTimeOffsetMapper.Instance);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the offset from UTC given to dates without one: cells that Excel stores as dates
    /// (which have no time zone), and text without an offset. By default they're UTC. Text with
    /// its own offset keeps it.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="offset">The offset, a whole number of minutes between -14 and +14 hours.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithOffset<TMap>(this ValuePipeline<DateTimeOffset?, TMap> map, TimeSpan offset) where TMap : ValuePipeline<DateTimeOffset?, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.Offset = offset;
        map.Pipeline.AddMapperIfMissing(DateTimeOffsetMapper.Instance);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a DateTimeOffset. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a DateTimeOffset.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormats<TMap>(this ValuePipeline<DateTimeOffset, TMap> map, params string[] formats) where TMap : ValuePipeline<DateTimeOffset, TMap>
    {
        map.Pipeline.SetFormats(formats);
        map.Pipeline.AddMapperIfMissing(DateTimeOffsetMapper.Instance);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a DateTimeOffset. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a DateTimeOffset.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormats<TMap>(this ValuePipeline<DateTimeOffset, TMap> map, IEnumerable<string> formats) where TMap : ValuePipeline<DateTimeOffset, TMap>
    {
        ThrowHelpers.ThrowIfNull(formats, nameof(formats));
        return map.WithFormats([.. formats]);
    }

    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a DateTimeOffset. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a DateTimeOffset.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormats<TMap>(this ValuePipeline<DateTimeOffset?, TMap> map, params string[] formats) where TMap : ValuePipeline<DateTimeOffset?, TMap>
    {
        map.Pipeline.SetFormats(formats);
        map.Pipeline.AddMapperIfMissing(DateTimeOffsetMapper.Instance);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a DateTimeOffset. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a DateTimeOffset.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormats<TMap>(this ValuePipeline<DateTimeOffset?, TMap> map, IEnumerable<string> formats) where TMap : ValuePipeline<DateTimeOffset?, TMap>
    {
        ThrowHelpers.ThrowIfNull(formats, nameof(formats));
        return map.WithFormats([.. formats]);
    }

    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a TimeSpan. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a TimeSpan.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormats<TMap>(this ValuePipeline<TimeSpan, TMap> map, params string[] formats) where TMap : ValuePipeline<TimeSpan, TMap>
    {
        map.Pipeline.SetFormats(formats);
        map.Pipeline.AddMapperIfMissing(TimeSpanMapper.Instance);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a TimeSpan. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a TimeSpan.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormats<TMap>(this ValuePipeline<TimeSpan, TMap> map, IEnumerable<string> formats) where TMap : ValuePipeline<TimeSpan, TMap>
    {
        ThrowHelpers.ThrowIfNull(formats, nameof(formats));
        return map.WithFormats([.. formats]);
    }

    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a TimeSpan. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a TimeSpan.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormats<TMap>(this ValuePipeline<TimeSpan?, TMap> map, params string[] formats) where TMap : ValuePipeline<TimeSpan?, TMap>
    {
        map.Pipeline.SetFormats(formats);
        map.Pipeline.AddMapperIfMissing(TimeSpanMapper.Instance);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a TimeSpan. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a TimeSpan.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormats<TMap>(this ValuePipeline<TimeSpan?, TMap> map, IEnumerable<string> formats) where TMap : ValuePipeline<TimeSpan?, TMap>
    {
        ThrowHelpers.ThrowIfNull(formats, nameof(formats));
        return map.WithFormats([.. formats]);
    }

#if NET6_0_OR_GREATER
    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a DateOnly. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a DateOnly.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormats<TMap>(this ValuePipeline<DateOnly, TMap> map, params string[] formats) where TMap : ValuePipeline<DateOnly, TMap>
    {
        map.Pipeline.SetFormats(formats);
        map.Pipeline.AddMapperIfMissing(DateOnlyMapper.Instance);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a DateOnly. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a DateOnly.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormats<TMap>(this ValuePipeline<DateOnly, TMap> map, IEnumerable<string> formats) where TMap : ValuePipeline<DateOnly, TMap>
    {
        ThrowHelpers.ThrowIfNull(formats, nameof(formats));
        return map.WithFormats([.. formats]);
    }

    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a DateOnly. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a DateOnly.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormats<TMap>(this ValuePipeline<DateOnly?, TMap> map, params string[] formats) where TMap : ValuePipeline<DateOnly?, TMap>
    {
        map.Pipeline.SetFormats(formats);
        map.Pipeline.AddMapperIfMissing(DateOnlyMapper.Instance);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a DateOnly. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a DateOnly.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormats<TMap>(this ValuePipeline<DateOnly?, TMap> map, IEnumerable<string> formats) where TMap : ValuePipeline<DateOnly?, TMap>
    {
        ThrowHelpers.ThrowIfNull(formats, nameof(formats));
        return map.WithFormats([.. formats]);
    }

    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a TimeOnly. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a TimeOnly.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormats<TMap>(this ValuePipeline<TimeOnly, TMap> map, params string[] formats) where TMap : ValuePipeline<TimeOnly, TMap>
    {
        map.Pipeline.SetFormats(formats);
        map.Pipeline.AddMapperIfMissing(TimeOnlyMapper.Instance);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a TimeOnly. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a TimeOnly.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormats<TMap>(this ValuePipeline<TimeOnly, TMap> map, IEnumerable<string> formats) where TMap : ValuePipeline<TimeOnly, TMap>
    {
        ThrowHelpers.ThrowIfNull(formats, nameof(formats));
        return map.WithFormats([.. formats]);
    }

    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a TimeOnly. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a TimeOnly.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormats<TMap>(this ValuePipeline<TimeOnly?, TMap> map, params string[] formats) where TMap : ValuePipeline<TimeOnly?, TMap>
    {
        map.Pipeline.SetFormats(formats);
        map.Pipeline.AddMapperIfMissing(TimeOnlyMapper.Instance);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies formats used when mapping the value of a cell to a TimeOnly. This is useful for
    /// mapping columns where formats differ. Existing formats are overriden.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">A list of formats to use when mapping the value of a cell to a TimeOnly.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormats<TMap>(this ValuePipeline<TimeOnly?, TMap> map, IEnumerable<string> formats) where TMap : ValuePipeline<TimeOnly?, TMap>
    {
        ThrowHelpers.ThrowIfNull(formats, nameof(formats));
        return map.WithFormats([.. formats]);
    }
#endif

    // A converter (WithConverter) replaces the built-in conversion, so setting the options of the
    // built-in conversion after one adds it back, after the converter.
    private static void AddMapperIfMissing<T>(this CellPipeline<T> pipeline, ICellMapper<T> mapper)
    {
        if (!pipeline.Mappers.Contains(mapper))
        {
            pipeline.Mappers.Add(mapper);
        }
    }

    private static void AddMapperIfMissing<T, TCellMapper>(this CellPipeline<T> pipeline, Func<TCellMapper> createMapper) where TCellMapper : ICellMapper<T>
    {
        foreach (var existing in pipeline.Mappers)
        {
            if (existing is TCellMapper)
            {
                return;
            }
        }

        pipeline.Mappers.Add(createMapper());
    }

    private static void SetFormats(this CellPipeline pipeline, string[] formats)
    {
        FormatUtilities.ValidateFormats(formats);
        pipeline.Options.Formats = formats;
    }

    /// <summary>
    /// Specifies the format provider used when mapping the value of a cell to a DateTime. This is useful for
    /// mapping columns where cultures differ. Existing provider is overridden.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider to use when mapping the value of a cell to a DateTime.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormatProvider<TMap>(this ValuePipeline<DateTime, TMap> map, IFormatProvider? provider) where TMap : ValuePipeline<DateTime, TMap>
    {
        map.Pipeline.Options.Provider = provider;
        map.Pipeline.AddMapperIfMissing(DateTimeMapper.Instance);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the format provider used when mapping the value of a cell to a DateTime. This is useful for
    /// mapping columns where cultures differ. Existing provider is overridden.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider to use when mapping the value of a cell to a DateTime.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormatProvider<TMap>(this ValuePipeline<DateTime?, TMap> map, IFormatProvider? provider) where TMap : ValuePipeline<DateTime?, TMap>
    {
        map.Pipeline.Options.Provider = provider;
        map.Pipeline.AddMapperIfMissing(DateTimeMapper.Instance);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the format provider used when mapping the value of a cell to a DateTimeOffset. This is useful for
    /// mapping columns where cultures differ. Existing provider is overridden.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider to use when mapping the value of a cell to a DateTimeOffset.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormatProvider<TMap>(this ValuePipeline<DateTimeOffset, TMap> map, IFormatProvider? provider) where TMap : ValuePipeline<DateTimeOffset, TMap>
    {
        map.Pipeline.Options.Provider = provider;
        map.Pipeline.AddMapperIfMissing(DateTimeOffsetMapper.Instance);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the format provider used when mapping the value of a cell to a DateTimeOffset. This is useful for
    /// mapping columns where cultures differ. Existing provider is overridden.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider to use when mapping the value of a cell to a DateTimeOffset.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormatProvider<TMap>(this ValuePipeline<DateTimeOffset?, TMap> map, IFormatProvider? provider) where TMap : ValuePipeline<DateTimeOffset?, TMap>
    {
        map.Pipeline.Options.Provider = provider;
        map.Pipeline.AddMapperIfMissing(DateTimeOffsetMapper.Instance);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the format provider used when mapping the value of a cell to a TimeSpan. This is useful for
    /// mapping columns where cultures differ. Existing provider is overridden.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider to use when mapping the value of a cell to a TimeSpan.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormatProvider<TMap>(this ValuePipeline<TimeSpan, TMap> map, IFormatProvider? provider) where TMap : ValuePipeline<TimeSpan, TMap>
    {
        map.Pipeline.Options.Provider = provider;
        map.Pipeline.AddMapperIfMissing(TimeSpanMapper.Instance);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the format provider used when mapping the value of a cell to a TimeSpan. This is useful for
    /// mapping columns where cultures differ. Existing provider is overridden.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider to use when mapping the value of a cell to a TimeSpan.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormatProvider<TMap>(this ValuePipeline<TimeSpan?, TMap> map, IFormatProvider? provider) where TMap : ValuePipeline<TimeSpan?, TMap>
    {
        map.Pipeline.Options.Provider = provider;
        map.Pipeline.AddMapperIfMissing(TimeSpanMapper.Instance);
        return (TMap)map;
    }

#if NET6_0_OR_GREATER
    /// <summary>
    /// Specifies the format provider used when mapping the value of a cell to a DateOnly. This is useful for
    /// mapping columns where cultures differ. Existing provider is overridden.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider to use when mapping the value of a cell to a DateOnly.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormatProvider<TMap>(this ValuePipeline<DateOnly, TMap> map, IFormatProvider? provider) where TMap : ValuePipeline<DateOnly, TMap>
    {
        map.Pipeline.Options.Provider = provider;
        map.Pipeline.AddMapperIfMissing(DateOnlyMapper.Instance);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the format provider used when mapping the value of a cell to a DateOnly. This is useful for
    /// mapping columns where cultures differ. Existing provider is overridden.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider to use when mapping the value of a cell to a DateOnly.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormatProvider<TMap>(this ValuePipeline<DateOnly?, TMap> map, IFormatProvider? provider) where TMap : ValuePipeline<DateOnly?, TMap>
    {
        map.Pipeline.Options.Provider = provider;
        map.Pipeline.AddMapperIfMissing(DateOnlyMapper.Instance);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the format provider used when mapping the value of a cell to a TimeOnly. This is useful for
    /// mapping columns where cultures differ. Existing provider is overridden.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider to use when mapping the value of a cell to a TimeOnly.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormatProvider<TMap>(this ValuePipeline<TimeOnly, TMap> map, IFormatProvider? provider) where TMap : ValuePipeline<TimeOnly, TMap>
    {
        map.Pipeline.Options.Provider = provider;
        map.Pipeline.AddMapperIfMissing(TimeOnlyMapper.Instance);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the format provider used when mapping the value of a cell to a TimeOnly. This is useful for
    /// mapping columns where cultures differ. Existing provider is overridden.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider to use when mapping the value of a cell to a TimeOnly.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormatProvider<TMap>(this ValuePipeline<TimeOnly?, TMap> map, IFormatProvider? provider) where TMap : ValuePipeline<TimeOnly?, TMap>
    {
        map.Pipeline.Options.Provider = provider;
        map.Pipeline.AddMapperIfMissing(TimeOnlyMapper.Instance);
        return (TMap)map;
    }
#endif

#if NET7_0_OR_GREATER
    /// <summary>
    /// Specifies the format provider used when mapping the value of a cell to a numeric type
    /// that implements <see cref="INumberBase{T}"/>. This is useful for mapping columns where
    /// cultures differ. Existing provider is overridden.
    /// </summary>
    /// <typeparam name="T">The type of the property or field that the map represents.</typeparam>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider to use when mapping the value of a cell to a numeric type.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormatProvider<T, TMap>(this ValuePipeline<T, TMap> map, IFormatProvider? provider) where T : INumberBase<T> where TMap : ValuePipeline<T, TMap>
    {
        map.Pipeline.Options.Provider = provider;
        map.Pipeline.AddMapperIfMissing<T, INumberBaseMapper<T>>(static () => new INumberBaseMapper<T>());
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the format provider used when mapping the value of a cell to a numeric type
    /// that implements <see cref="INumberBase{T}"/>. This is useful for mapping columns where
    /// cultures differ. Existing provider is overridden.
    /// </summary>
    /// <typeparam name="T">The type of the property or field that the map represents.</typeparam>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider to use when mapping the value of a cell to a numeric type.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormatProvider<T, TMap>(this ValuePipeline<T?, TMap> map, IFormatProvider? provider) where T : struct, INumberBase<T> where TMap : ValuePipeline<T?, TMap>
    {
        map.Pipeline.Options.Provider = provider;
        map.Pipeline.AddMapperIfMissing<T?, StructNumberMapper<T>>(static () => new StructNumberMapper<T>());
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the number styles to use when mapping the value of a cell to a numeric type
    /// that implements <see cref="INumberBase{T}"/>. This is useful for specifying
    /// how numeric values are formatted in the Excel file.
    /// </summary>
    /// <typeparam name="T">The type of the property or field that the map represents.</typeparam>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="style">The number styles to use when mapping the value of a cell to a numeric type.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithNumberStyle<T, TMap>(this ValuePipeline<T, TMap> map, NumberStyles style) where T : INumberBase<T> where TMap : ValuePipeline<T, TMap>
    {
        map.Pipeline.Options.NumberStyle = style;
        map.Pipeline.AddMapperIfMissing<T, INumberBaseMapper<T>>(static () => new INumberBaseMapper<T>());
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the number styles to use when mapping the value of a cell to a numeric type
    /// that implements <see cref="INumberBase{T}"/>. This is useful for specifying
    /// how numeric values are formatted in the Excel file.
    /// </summary>
    /// <typeparam name="T">The type of the property or field that the map represents.</typeparam>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="style">The number styles to use when mapping the value of a cell to a numeric type.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithNumberStyle<T, TMap>(this ValuePipeline<T?, TMap> map, NumberStyles style) where T : struct, INumberBase<T> where TMap : ValuePipeline<T?, TMap>
    {
        map.Pipeline.Options.NumberStyle = style;
        map.Pipeline.AddMapperIfMissing<T?, StructNumberMapper<T>>(static () => new StructNumberMapper<T>());
        return (TMap)map;
    }
#endif

    /// <summary>
    /// Matches the enum's names ignoring case, so that "active" maps to <c>Status.Active</c>.
    /// </summary>
    /// <typeparam name="TEnum">The enum type of the member.</typeparam>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithIgnoreCase<TEnum, TMap>(this ValuePipeline<TEnum, TMap> map) where TEnum : struct, Enum where TMap : ValuePipeline<TEnum, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.IgnoreCase = true;
        map.Pipeline.AddMapperIfMissing<TEnum, EnumMapper<TEnum>>(static () => new EnumMapper<TEnum>());
        return (TMap)map;
    }

    /// <summary>
    /// Matches the enum's names ignoring case, so that "active" maps to <c>Status.Active</c>.
    /// </summary>
    /// <typeparam name="TEnum">The enum type of the member.</typeparam>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithIgnoreCase<TEnum, TMap>(this ValuePipeline<TEnum?, TMap> map) where TEnum : struct, Enum where TMap : ValuePipeline<TEnum?, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.IgnoreCase = true;
        map.Pipeline.AddMapperIfMissing<TEnum?, EnumMapper<TEnum>>(static () => new EnumMapper<TEnum>());
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the kind of URI to use when mapping the value of a cell to a Uri. This is useful for
    /// specifying whether URIs should be absolute, relative, or either.
    /// </summary>
    /// <typeparam name="TUri">The type of the member: <see cref="Uri"/>.</typeparam>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="uriKind">The kind of URI to use when mapping the value of a cell to a Uri. Must be a valid UriKind value (Absolute, Relative, or RelativeOrAbsolute).</param>
    /// <returns>The map on which this method was invoked.</returns>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="uriKind"/> is not a valid UriKind value.</exception>
    public static TMap WithUriKind<TUri, TMap>(this ValuePipeline<TUri, TMap> map, UriKind uriKind) where TUri : Uri? where TMap : ValuePipeline<TUri, TMap>
    {
        EnumUtilities.ValidateIsDefined(uriKind);
        map.Pipeline.Options.UriKind = uriKind;
        map.Pipeline.AddMapperIfMissing((ICellMapper<TUri>)(object)UriMapper.Instance);
        return (TMap)map;
    }

    /// <summary>
    /// Converts cells with <paramref name="converter"/>. The converter replaces the built-in conversion
    /// for the member's type and any earlier converter, so it is used for every cell that isn't empty.
    /// Empty cells still use the empty fallback, and values the converter returns
    /// <see cref="CellConversionResult{T}.Invalid()"/> for use the invalid fallback.
    /// </summary>
    /// <typeparam name="T">The type of the member, element or value.</typeparam>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map.</param>
    /// <param name="converter">The converter.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithConverter<T, TMap>(this ValuePipeline<T, TMap> map, ICellConverter<T> converter) where TMap : ValuePipeline<T, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        ThrowHelpers.ThrowIfNull(converter, nameof(converter));

        map.Pipeline.Mappers.Clear();
        map.Pipeline.Mappers.Add(new CellConverterMapper<T>(converter));
        return (TMap)map;
    }

    /// <summary>
    /// Converts the cells of a nullable member with a converter for its underlying type, as for
    /// <see cref="WithConverter{T, TMap}(ValuePipeline{T, TMap}, ICellConverter{T})"/>. Empty cells
    /// still use the empty fallback, which is null by default.
    /// </summary>
    /// <typeparam name="T">The underlying type of the member, element or value.</typeparam>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map.</param>
    /// <param name="converter">The converter.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithConverter<T, TMap>(this ValuePipeline<T?, TMap> map, ICellConverter<T> converter) where T : struct where TMap : ValuePipeline<T?, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        ThrowHelpers.ThrowIfNull(converter, nameof(converter));

        map.Pipeline.Mappers.Clear();
        map.Pipeline.Mappers.Add(new CellConverterMapper<T?>(new NullableCellConverter<T>(converter)));
        return (TMap)map;
    }

    /// <summary>
    /// Gives empty cells the value <paramref name="fallback"/> returns.
    /// </summary>
    /// <typeparam name="T">The type of the member, element or value.</typeparam>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map.</param>
    /// <param name="fallback">The fallback.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithEmptyFallback<T, TMap>(this ValuePipeline<T, TMap> map, ICellFallback<T> fallback) where TMap : ValuePipeline<T, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        ThrowHelpers.ThrowIfNull(fallback, nameof(fallback));
        map.Pipeline.EmptyFallback = fallback;
        return (TMap)map;
    }

    /// <summary>
    /// Gives empty cells of a nullable member the value a fallback for its underlying type returns.
    /// </summary>
    /// <typeparam name="T">The underlying type of the member, element or value.</typeparam>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map.</param>
    /// <param name="fallback">The fallback.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithEmptyFallback<T, TMap>(this ValuePipeline<T?, TMap> map, ICellFallback<T> fallback) where T : struct where TMap : ValuePipeline<T?, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        ThrowHelpers.ThrowIfNull(fallback, nameof(fallback));
        map.Pipeline.EmptyFallback = new NullableCellFallback<T>(fallback);
        return (TMap)map;
    }

    /// <summary>
    /// Gives invalid cells the value <paramref name="fallback"/> returns.
    /// </summary>
    /// <typeparam name="T">The type of the member, element or value.</typeparam>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map.</param>
    /// <param name="fallback">The fallback.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithInvalidFallback<T, TMap>(this ValuePipeline<T, TMap> map, ICellFallback<T> fallback) where TMap : ValuePipeline<T, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        ThrowHelpers.ThrowIfNull(fallback, nameof(fallback));
        map.Pipeline.InvalidFallback = fallback;
        return (TMap)map;
    }

    /// <summary>
    /// Gives invalid cells of a nullable member the value a fallback for its underlying type returns.
    /// </summary>
    /// <typeparam name="T">The underlying type of the member, element or value.</typeparam>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map.</param>
    /// <param name="fallback">The fallback.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithInvalidFallback<T, TMap>(this ValuePipeline<T?, TMap> map, ICellFallback<T> fallback) where T : struct where TMap : ValuePipeline<T?, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        ThrowHelpers.ThrowIfNull(fallback, nameof(fallback));
        map.Pipeline.InvalidFallback = new NullableCellFallback<T>(fallback);
        return (TMap)map;
    }
}

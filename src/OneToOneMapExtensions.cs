using System.Globalization;
#if NET7_0_OR_GREATER
using System.Numerics;
#endif
using ExcelMapper.Fallbacks;
using ExcelMapper.Mappers;

namespace ExcelMapper;

/// <summary>
/// Methods for <see cref="OneToOneMap{T}"/> of particular types, which return the map itself, so that
/// calls chain, for example <c>Map(e =&gt; e.Joined).WithFormats("yyyy-MM-dd").MakeOptional()</c>.
/// </summary>
public static class OneToOneMapExtensions
{
    /// <summary>
    /// Specifies the formats used to parse text as a DateTime. Existing formats are replaced.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">The formats.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static OneToOneMap<DateTime> WithFormats(this OneToOneMap<DateTime> map, params string[] formats)
    {
        ((ValuePipeline<DateTime>)map).WithFormats(formats);
        return map;
    }

    /// <summary>
    /// Specifies the formats used to parse text as a DateTime. Existing formats are replaced.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">The formats.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static OneToOneMap<DateTime> WithFormats(this OneToOneMap<DateTime> map, params IEnumerable<string> formats)
    {
        ((ValuePipeline<DateTime>)map).WithFormats(formats);
        return map;
    }

    /// <summary>
    /// Specifies the format provider used to parse text as a DateTime.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static OneToOneMap<DateTime> WithFormatProvider(this OneToOneMap<DateTime> map, IFormatProvider? provider)
    {
        ((ValuePipeline<DateTime>)map).WithFormatProvider(provider);
        return map;
    }

    /// <summary>
    /// Specifies the formats used to parse text as a DateTime. Existing formats are replaced.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">The formats.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static OneToOneMap<DateTime?> WithFormats(this OneToOneMap<DateTime?> map, params string[] formats)
    {
        ((ValuePipeline<DateTime?>)map).WithFormats(formats);
        return map;
    }

    /// <summary>
    /// Specifies the formats used to parse text as a DateTime. Existing formats are replaced.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">The formats.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static OneToOneMap<DateTime?> WithFormats(this OneToOneMap<DateTime?> map, params IEnumerable<string> formats)
    {
        ((ValuePipeline<DateTime?>)map).WithFormats(formats);
        return map;
    }

    /// <summary>
    /// Specifies the format provider used to parse text as a DateTime.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static OneToOneMap<DateTime?> WithFormatProvider(this OneToOneMap<DateTime?> map, IFormatProvider? provider)
    {
        ((ValuePipeline<DateTime?>)map).WithFormatProvider(provider);
        return map;
    }

    /// <summary>
    /// Specifies the formats used to parse text as a DateTimeOffset. Existing formats are replaced.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">The formats.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static OneToOneMap<DateTimeOffset> WithFormats(this OneToOneMap<DateTimeOffset> map, params string[] formats)
    {
        ((ValuePipeline<DateTimeOffset>)map).WithFormats(formats);
        return map;
    }

    /// <summary>
    /// Specifies the formats used to parse text as a DateTimeOffset. Existing formats are replaced.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">The formats.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static OneToOneMap<DateTimeOffset> WithFormats(this OneToOneMap<DateTimeOffset> map, params IEnumerable<string> formats)
    {
        ((ValuePipeline<DateTimeOffset>)map).WithFormats(formats);
        return map;
    }

    /// <summary>
    /// Specifies the format provider used to parse text as a DateTimeOffset.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static OneToOneMap<DateTimeOffset> WithFormatProvider(this OneToOneMap<DateTimeOffset> map, IFormatProvider? provider)
    {
        ((ValuePipeline<DateTimeOffset>)map).WithFormatProvider(provider);
        return map;
    }

    /// <summary>
    /// Specifies the offset from UTC given to cells that Excel stores as dates.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="offset">The offset.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static OneToOneMap<DateTimeOffset> WithOffset(this OneToOneMap<DateTimeOffset> map, TimeSpan offset)
    {
        ((ValuePipeline<DateTimeOffset>)map).WithOffset(offset);
        return map;
    }

    /// <summary>
    /// Specifies the formats used to parse text as a DateTimeOffset. Existing formats are replaced.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">The formats.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static OneToOneMap<DateTimeOffset?> WithFormats(this OneToOneMap<DateTimeOffset?> map, params string[] formats)
    {
        ((ValuePipeline<DateTimeOffset?>)map).WithFormats(formats);
        return map;
    }

    /// <summary>
    /// Specifies the formats used to parse text as a DateTimeOffset. Existing formats are replaced.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">The formats.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static OneToOneMap<DateTimeOffset?> WithFormats(this OneToOneMap<DateTimeOffset?> map, params IEnumerable<string> formats)
    {
        ((ValuePipeline<DateTimeOffset?>)map).WithFormats(formats);
        return map;
    }

    /// <summary>
    /// Specifies the format provider used to parse text as a DateTimeOffset.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static OneToOneMap<DateTimeOffset?> WithFormatProvider(this OneToOneMap<DateTimeOffset?> map, IFormatProvider? provider)
    {
        ((ValuePipeline<DateTimeOffset?>)map).WithFormatProvider(provider);
        return map;
    }

    /// <summary>
    /// Specifies the offset from UTC given to cells that Excel stores as dates.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="offset">The offset.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static OneToOneMap<DateTimeOffset?> WithOffset(this OneToOneMap<DateTimeOffset?> map, TimeSpan offset)
    {
        ((ValuePipeline<DateTimeOffset?>)map).WithOffset(offset);
        return map;
    }

    /// <summary>
    /// Specifies the formats used to parse text as a TimeSpan. Existing formats are replaced.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">The formats.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static OneToOneMap<TimeSpan> WithFormats(this OneToOneMap<TimeSpan> map, params string[] formats)
    {
        ((ValuePipeline<TimeSpan>)map).WithFormats(formats);
        return map;
    }

    /// <summary>
    /// Specifies the formats used to parse text as a TimeSpan. Existing formats are replaced.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">The formats.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static OneToOneMap<TimeSpan> WithFormats(this OneToOneMap<TimeSpan> map, params IEnumerable<string> formats)
    {
        ((ValuePipeline<TimeSpan>)map).WithFormats(formats);
        return map;
    }

    /// <summary>
    /// Specifies the format provider used to parse text as a TimeSpan.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static OneToOneMap<TimeSpan> WithFormatProvider(this OneToOneMap<TimeSpan> map, IFormatProvider? provider)
    {
        ((ValuePipeline<TimeSpan>)map).WithFormatProvider(provider);
        return map;
    }

    /// <summary>
    /// Specifies the formats used to parse text as a TimeSpan. Existing formats are replaced.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">The formats.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static OneToOneMap<TimeSpan?> WithFormats(this OneToOneMap<TimeSpan?> map, params string[] formats)
    {
        ((ValuePipeline<TimeSpan?>)map).WithFormats(formats);
        return map;
    }

    /// <summary>
    /// Specifies the formats used to parse text as a TimeSpan. Existing formats are replaced.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">The formats.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static OneToOneMap<TimeSpan?> WithFormats(this OneToOneMap<TimeSpan?> map, params IEnumerable<string> formats)
    {
        ((ValuePipeline<TimeSpan?>)map).WithFormats(formats);
        return map;
    }

    /// <summary>
    /// Specifies the format provider used to parse text as a TimeSpan.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static OneToOneMap<TimeSpan?> WithFormatProvider(this OneToOneMap<TimeSpan?> map, IFormatProvider? provider)
    {
        ((ValuePipeline<TimeSpan?>)map).WithFormatProvider(provider);
        return map;
    }

#if NET6_0_OR_GREATER
    /// <summary>
    /// Specifies the formats used to parse text as a DateOnly. Existing formats are replaced.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">The formats.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static OneToOneMap<DateOnly> WithFormats(this OneToOneMap<DateOnly> map, params string[] formats)
    {
        ((ValuePipeline<DateOnly>)map).WithFormats(formats);
        return map;
    }

    /// <summary>
    /// Specifies the formats used to parse text as a DateOnly. Existing formats are replaced.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">The formats.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static OneToOneMap<DateOnly> WithFormats(this OneToOneMap<DateOnly> map, params IEnumerable<string> formats)
    {
        ((ValuePipeline<DateOnly>)map).WithFormats(formats);
        return map;
    }

    /// <summary>
    /// Specifies the format provider used to parse text as a DateOnly.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static OneToOneMap<DateOnly> WithFormatProvider(this OneToOneMap<DateOnly> map, IFormatProvider? provider)
    {
        ((ValuePipeline<DateOnly>)map).WithFormatProvider(provider);
        return map;
    }

    /// <summary>
    /// Specifies the formats used to parse text as a DateOnly. Existing formats are replaced.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">The formats.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static OneToOneMap<DateOnly?> WithFormats(this OneToOneMap<DateOnly?> map, params string[] formats)
    {
        ((ValuePipeline<DateOnly?>)map).WithFormats(formats);
        return map;
    }

    /// <summary>
    /// Specifies the formats used to parse text as a DateOnly. Existing formats are replaced.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">The formats.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static OneToOneMap<DateOnly?> WithFormats(this OneToOneMap<DateOnly?> map, params IEnumerable<string> formats)
    {
        ((ValuePipeline<DateOnly?>)map).WithFormats(formats);
        return map;
    }

    /// <summary>
    /// Specifies the format provider used to parse text as a DateOnly.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static OneToOneMap<DateOnly?> WithFormatProvider(this OneToOneMap<DateOnly?> map, IFormatProvider? provider)
    {
        ((ValuePipeline<DateOnly?>)map).WithFormatProvider(provider);
        return map;
    }
#endif
#if NET6_0_OR_GREATER
    /// <summary>
    /// Specifies the formats used to parse text as a TimeOnly. Existing formats are replaced.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">The formats.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static OneToOneMap<TimeOnly> WithFormats(this OneToOneMap<TimeOnly> map, params string[] formats)
    {
        ((ValuePipeline<TimeOnly>)map).WithFormats(formats);
        return map;
    }

    /// <summary>
    /// Specifies the formats used to parse text as a TimeOnly. Existing formats are replaced.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">The formats.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static OneToOneMap<TimeOnly> WithFormats(this OneToOneMap<TimeOnly> map, params IEnumerable<string> formats)
    {
        ((ValuePipeline<TimeOnly>)map).WithFormats(formats);
        return map;
    }

    /// <summary>
    /// Specifies the format provider used to parse text as a TimeOnly.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static OneToOneMap<TimeOnly> WithFormatProvider(this OneToOneMap<TimeOnly> map, IFormatProvider? provider)
    {
        ((ValuePipeline<TimeOnly>)map).WithFormatProvider(provider);
        return map;
    }

    /// <summary>
    /// Specifies the formats used to parse text as a TimeOnly. Existing formats are replaced.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">The formats.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static OneToOneMap<TimeOnly?> WithFormats(this OneToOneMap<TimeOnly?> map, params string[] formats)
    {
        ((ValuePipeline<TimeOnly?>)map).WithFormats(formats);
        return map;
    }

    /// <summary>
    /// Specifies the formats used to parse text as a TimeOnly. Existing formats are replaced.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="formats">The formats.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static OneToOneMap<TimeOnly?> WithFormats(this OneToOneMap<TimeOnly?> map, params IEnumerable<string> formats)
    {
        ((ValuePipeline<TimeOnly?>)map).WithFormats(formats);
        return map;
    }

    /// <summary>
    /// Specifies the format provider used to parse text as a TimeOnly.
    /// </summary>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static OneToOneMap<TimeOnly?> WithFormatProvider(this OneToOneMap<TimeOnly?> map, IFormatProvider? provider)
    {
        ((ValuePipeline<TimeOnly?>)map).WithFormatProvider(provider);
        return map;
    }
#endif
#if NET7_0_OR_GREATER
    /// <summary>
    /// Specifies the format provider used to parse text as a number.
    /// </summary>
    /// <typeparam name="T">The type of the property or field that the map represents.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static OneToOneMap<T> WithFormatProvider<T>(this OneToOneMap<T> map, IFormatProvider? provider) where T : INumberBase<T>
    {
        ((ValuePipeline<T>)map).WithFormatProvider(provider);
        return map;
    }

    /// <summary>
    /// Specifies the format provider used to parse text as a number.
    /// </summary>
    /// <typeparam name="T">The underlying type of the property or field that the map represents.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static OneToOneMap<T?> WithFormatProvider<T>(this OneToOneMap<T?> map, IFormatProvider? provider) where T : struct, INumberBase<T>
    {
        ((ValuePipeline<T?>)map).WithFormatProvider(provider);
        return map;
    }

    /// <summary>
    /// Specifies the number styles used to parse text as a number.
    /// </summary>
    /// <typeparam name="T">The type of the property or field that the map represents.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="style">The number styles.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static OneToOneMap<T> WithNumberStyle<T>(this OneToOneMap<T> map, NumberStyles style) where T : INumberBase<T>
    {
        ((ValuePipeline<T>)map).WithNumberStyle(style);
        return map;
    }

    /// <summary>
    /// Specifies the number styles used to parse text as a number.
    /// </summary>
    /// <typeparam name="T">The underlying type of the property or field that the map represents.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="style">The number styles.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static OneToOneMap<T?> WithNumberStyle<T>(this OneToOneMap<T?> map, NumberStyles style) where T : struct, INumberBase<T>
    {
        ((ValuePipeline<T?>)map).WithNumberStyle(style);
        return map;
    }
#endif

    /// <summary>
    /// Converts the cells of a nullable member with a converter for its underlying type. The converter
    /// replaces the built-in conversion; empty cells still use the empty fallback, which is null by default.
    /// </summary>
    /// <typeparam name="T">The underlying type of the property or field that the map represents.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="converter">The converter.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static OneToOneMap<T?> WithConverter<T>(this OneToOneMap<T?> map, ICellConverter<T> converter) where T : struct
    {
        ((ValuePipeline<T?>)map).WithConverter(converter);
        return map;
    }

    /// <summary>
    /// Gives empty cells of a nullable member the value a fallback for its underlying type returns.
    /// </summary>
    /// <typeparam name="T">The underlying type of the property or field that the map represents.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="fallback">The fallback.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static OneToOneMap<T?> WithEmptyFallback<T>(this OneToOneMap<T?> map, ICellFallback<T> fallback) where T : struct
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        ThrowHelpers.ThrowIfNull(fallback, nameof(fallback));
        map.EmptyFallback = new NullableCellFallback<T>(fallback);
        return map;
    }

    /// <summary>
    /// Gives invalid cells of a nullable member the value a fallback for its underlying type returns.
    /// </summary>
    /// <typeparam name="T">The underlying type of the property or field that the map represents.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="fallback">The fallback.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static OneToOneMap<T?> WithInvalidFallback<T>(this OneToOneMap<T?> map, ICellFallback<T> fallback) where T : struct
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        ThrowHelpers.ThrowIfNull(fallback, nameof(fallback));
        map.InvalidFallback = new NullableCellFallback<T>(fallback);
        return map;
    }
}

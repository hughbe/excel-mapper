using ExcelMapper.Abstractions;
using ExcelMapper.Fallbacks;
using ExcelMapper.Mappers;

namespace ExcelMapper;
 
public delegate T ConvertUsingSimpleMapperDelegate<out T>(object value);

/// <summary>
/// Extensions on IMapper to enable fluent "With" method chaining.
/// </summary>
public static class IMapperExtensions
{
    /// <summary>
    /// Specifies that the string value of the cell should be trimmed before it is mapped to
    /// a property or field.
    /// </summary>
    /// <typeparam name="T">The type of the property map.</typeparam>
    /// <param name="propertyMap">The property map to use.</param>
    /// <returns>The property map on which this method was invoked.</returns>
    public static T WithTrim<T>(this T propertyMap) where T : IMapper
    {
        propertyMap.Mappers.Insert(0, new TrimCellValueTransformer());
        return propertyMap;
    }

    /// <summary>
    /// Specifies additional custom mappers that will be used to map the value of a cell to
    /// a property or field.
    /// </summary>
    /// <typeparam name="TPropertyMap">The type of the property map.</typeparam>
    /// <param name="propertyMap">The property map to use.</param>
    /// <param name="mappers">A list of additional custom mappers that will be used to map the value of a cell to a property or field</param>
    /// <returns>The property map on which this method was invoked.</returns>
    public static TPropertyMap WithCellValueMappers<TPropertyMap>(this TPropertyMap propertyMap, params ICellValueMapper[] mappers) where TPropertyMap : IMapper
    {
        if (mappers == null)
        {
            throw new ArgumentNullException(nameof(mappers));
        }

        foreach (ICellValueMapper mapper in mappers)
        {
            if (mapper == null)
            {
                throw new ArgumentNullException(nameof(mappers));
            }
        }

        foreach (ICellValueMapper mapper in mappers)
        {
            propertyMap.AppendMapper(mapper);
        }

        return propertyMap;
    }

    /// <summary>
    /// Specifies that the value of a cell should be mapped to a fixed value if it cannot be parsed. This
    /// is useful for mapping columns where equivilent data was entered differently.
    /// </summary>
    /// <typeparam name="TPropertyMap">The type of the property map.</typeparam>
    /// <typeparam name="T">The type of the property or field that the property map represents.</typeparam>
    /// <param name="propertyMap">The property map to use.</param>
    /// <param name="mappingDictionary">A dictionary that maps a fixed string value to a fixed value of T.</param>
    /// <param name="comparer">The comparer uses to map fixed string values. This allows for case-insensitive mappings, for example.</param>
    /// <returns>The property map on which this method was invoked.</returns>
    public static TPropertyMap WithMapping<TPropertyMap, T>(this TPropertyMap propertyMap, IDictionary<string, T> mappingDictionary, IEqualityComparer<string> comparer = null) where TPropertyMap : IMapper<T>
    {
        var item = new DictionaryMapper<T>(mappingDictionary, comparer);
        propertyMap.AppendMapper(item);
        return propertyMap;
    }

    /// <summary>
    /// Specifies data formats used when mapping the value of a cell to a DateTime. This is useful for
    /// mapping columns where data formats differ. Existing date formats are overriden.
    /// </summary>
    /// <param name="propertyMap">The property map to use.</param>
    /// <param name="formats">A list of date formats to use when mapping the value of a cell to a DateTime.</param>
    /// <returns>The property map on which this method was invoked.</returns>
    public static IMapper<DateTime> WithDateFormats(this IMapper<DateTime> propertyMap, params string[] formats)
    {
        propertyMap.AddFormats(formats);
        return propertyMap;
    }

    /// <summary>
    /// Specifies data formats used when mapping the value of a cell to a DateTime. This is useful for
    /// mapping columns where data formats differ. Existing date formats are overriden.
    /// </summary>
    /// <param name="propertyMap">The property map to use.</param>
    /// <param name="formats">A list of date formats to use when mapping the value of a cell to a DateTime.</param>
    /// <returns>The property map on which this method was invoked.</returns>
    public static IMapper<DateTime> WithDateFormats(this IMapper<DateTime> propertyMap, IEnumerable<string> formats)
    {
        return propertyMap.WithDateFormats(formats?.ToArray());
    }

    /// <summary>
    /// Specifies data formats used when mapping the value of a cell to a DateTime. This is useful for
    /// mapping columns where data formats differ. Existing date formats are overriden.
    /// </summary>
    /// <param name="propertyMap">The property map to use.</param>
    /// <param name="formats">A list of date formats to use when mapping the value of a cell to a DateTime.</param>
    /// <returns>The property map on which this method was invoked.</returns>
    public static IMapper<DateTime?> WithDateFormats(this IMapper<DateTime?> propertyMap, params string[] formats)
    {
        propertyMap.AddFormats(formats);
        return propertyMap;
    }

    /// <summary>
    /// Specifies data formats used when mapping the value of a cell to a DateTime. This is useful for
    /// mapping columns where data formats differ. Existing date formats are overriden.
    /// </summary>
    /// <param name="propertyMap">The property map to use.</param>
    /// <param name="formats">A list of date formats to use when mapping the value of a cell to a DateTime.</param>
    /// <returns>The property map on which this method was invoked.</returns>
    public static IMapper<DateTime?> WithDateFormats(this IMapper<DateTime?> propertyMap, IEnumerable<string> formats)
    {
        return propertyMap.WithDateFormats(formats?.ToArray());
    }

    private static void AddFormats(this IMapper propertyMap, string[] formats)
    {
        if (formats == null)
        {
            throw new ArgumentNullException(nameof(formats));
        }

        if (formats.Length == 0)
        {
            throw new ArgumentException("Formats cannot be empty.", nameof(formats));
        }

        DateTimeMapper dateTimeItem = (DateTimeMapper)propertyMap.Mappers.FirstOrDefault(item => item is DateTimeMapper);
        if (dateTimeItem == null)
        {
            dateTimeItem = new DateTimeMapper();
            propertyMap.AppendMapper(dateTimeItem);
        }

        dateTimeItem.Formats = formats;
    }

    /// <summary>
    /// Specifies that the value of a cell should be mapped to a value using the given delegate. This is
    /// useful for specifying custom mapping behaviour for a property or field without having to write
    /// your own ICellValueMapper.
    /// </summary>
    /// <typeparam name="TPropertyMap">The type of the property map.</typeparam>
    /// <typeparam name="T">The type of the property or field that the property map represents.</typeparam>
    /// <param name="propertyMap">The property map to use.</param>
    /// <param name="converter">A delegate that is invoked to map the string value of a cell to the value of a property or field.</param>
    /// <returns>The property map on which this method was invoked.</returns>
    public static TPropertyMap WithConverter<TPropertyMap, T>(this TPropertyMap propertyMap, ConvertUsingSimpleMapperDelegate<T> converter) where TPropertyMap : IMapper<T>
    {
        if (converter == null)
        {
            throw new ArgumentNullException(nameof(converter));
        }

        ConvertUsingMapperDelegate actualConverter = (ExcelCell cell, object value) =>
        {
            try
            {
                object result = converter(value);
                return CellValueMapperResult.Success(result);
            }
            catch (Exception exception)
            {
                return CellValueMapperResult.Invalid(exception);
            }
        };

        return propertyMap.WithConverter(actualConverter);
    }

    public static TPropertyMap WithConverter<TPropertyMap>(this TPropertyMap propertyMap, ConvertUsingMapperDelegate converter) where TPropertyMap: IMapper
    {
        if (converter == null)
        {
            throw new ArgumentNullException(nameof(converter));
        }

        var item = new ConvertUsingMapper(converter);
        propertyMap.AppendMapper(item);
        return propertyMap;
    }

    /// <summary>
    /// Specifies a fixed fallback to be used if the value of a cell is empty or cannot be mapped.
    /// </summary>
    /// <typeparam name="TPropertyMap">The type of the property map.</typeparam>
    /// <param name="propertyMap">The property map to use.</param>
    /// <param name="defaultValue">The value that will be assigned to the property or field if the value of a cell is empty or cannot be mapped.</param>
    /// <returns>The property map on which this method was invoked.</returns>
    public static TPropertyMap WithValueFallback<TPropertyMap>(this TPropertyMap propertyMap, object defaultValue) where TPropertyMap : IMapper
    {
        return propertyMap
            .WithEmptyFallback(defaultValue)
            .WithInvalidFallback(defaultValue);
    }

    /// <summary>
    /// Specifies that the property map should throw an exception if the value of a cell if empty or cannot be mapped.
    /// </summary>
    /// <typeparam name="TPropertyMap">The type of the property map.</typeparam>
    /// <param name="propertyMap">The property map to use.</param>
    /// <returns>The property map on which this method was invoked.</returns>
    public static TPropertyMap WithThrowingFallback<TPropertyMap>(this TPropertyMap propertyMap) where TPropertyMap : IMapper
    {
        return propertyMap
            .WithThrowingEmptyFallback()
            .WithThrowingInvalidFallback();
    }

    /// <summary>
    /// Specifies a fixed fallback to be used if the value of a cell is empty.
    /// </summary>
    /// <typeparam name="TPropertyMap">The type of the property map.</typeparam>
    /// <param name="propertyMap">The property map to use.</param>
    /// <param name="fallbackValue">The value that will be assigned to the property or field if the value of a cell is empty.</param>
    /// <returns>The property map on which this method was invoked.</returns>
    public static TPropertyMap WithEmptyFallback<TPropertyMap>(this TPropertyMap propertyMap, object fallbackValue) where TPropertyMap : IMapper
    {
        return propertyMap
            .WithEmptyFallbackItem(new FixedValueFallback(fallbackValue));
    }

    /// <summary>
    /// Specifies a custom fallback to be used if the value of a cell is empty.
    /// </summary>
    /// <typeparam name="TPropertyMap">The type of the property map.</typeparam>
    /// <param name="propertyMap">The property map to use.</param>
    /// <param name="fallbackItem">The fallback to be used if the value of a cell is empty.</param>
    /// <returns>The property map on which this method was invoked.</returns>
    public static TPropertyMap WithEmptyFallbackItem<TPropertyMap>(this TPropertyMap propertyMap, IEmptyCellFallback fallbackItem) where TPropertyMap : IMapper
    {
        if (fallbackItem == null)
        {
            throw new ArgumentNullException(nameof(fallbackItem));
        }

        if (propertyMap.Mappers.Count > 0 && propertyMap.Mappers[0] is EmptyCellMapper)
        {
            propertyMap.Mappers.RemoveAt(0);
        }

        propertyMap.Mappers.Insert(0, new EmptyCellMapper(fallbackItem));
        return propertyMap;
    }

    /// <summary>
    /// Specifies that the property map should throw an exception if the value of a cell is empty.
    /// </summary>
    /// <typeparam name="TPropertyMap">The type of the property map.</typeparam>
    /// <param name="propertyMap">The property map to use.</param>
    /// <returns>The property map on which this method was invoked.</returns>
    public static TPropertyMap WithThrowingEmptyFallback<TPropertyMap>(this TPropertyMap propertyMap) where TPropertyMap : IMapper
    {
        return propertyMap
            .WithEmptyFallbackItem(new ThrowFallback());
    }

    /// <summary>
    /// Specifies that the property map should throw an exception if the value of a cell cannot be mapped.
    /// </summary>
    /// <typeparam name="TPropertyMap">The type of the property map.</typeparam>
    /// <param name="propertyMap">The property map to use.</param>
    /// <returns>The property map on which this method was invoked.</returns>
    public static TPropertyMap WithThrowingInvalidFallback<TPropertyMap>(this TPropertyMap propertyMap) where TPropertyMap : IMapper
    {
        return propertyMap
            .WithInvalidFallbackItem(new ThrowFallback());
    }

    /// <summary>
    /// Specifies a fixed fallback to be used if the value of a cell cannot be mapped.
    /// </summary>
    /// <typeparam name="TPropertyMap">The type of the property map.</typeparam>
    /// <param name="propertyMap">The property map to use.</param>
    /// <param name="fallbackValue">The value that will be assigned to the property or field if the value of a cell cannot be mapped.</param>
    /// <returns>The property map on which this method was invoked.</returns>
    public static TPropertyMap WithInvalidFallback<TPropertyMap>(this TPropertyMap propertyMap, object fallbackValue) where TPropertyMap : IMapper
    {
        return propertyMap
            .WithInvalidFallbackItem(new FixedValueFallback(fallbackValue));
    }

    /// <summary>
    /// Specifies a custom fallback to be used if the value of a cell cannot be mapped.
    /// </summary>
    /// <typeparam name="TPropertyMap">The type of the property map.</typeparam>
    /// <param name="propertyMap">The property map to use.</param>
    /// <param name="fallbackItem">The fallback to be used if the value of a cell cannot be mapped.</param>
    /// <returns>The property map on which this method was invoked.</returns>
    public static TPropertyMap WithInvalidFallbackItem<TPropertyMap>(this TPropertyMap propertyMap, IInvalidCellFallback fallbackItem) where TPropertyMap : IMapper
    {
        if (fallbackItem == null)
        {
            throw new ArgumentNullException(nameof(fallbackItem));
        }

        if (propertyMap.Mappers.Count > 0 && propertyMap.Mappers[propertyMap.Mappers.Count - 1] is InvalidCellMapper)
        {
            propertyMap.Mappers.RemoveAt(propertyMap.Mappers.Count - 1);
        }

        propertyMap.AppendMapper(new InvalidCellMapper(fallbackItem));
        return propertyMap;
    }

    private static void AppendMapper<TPropertyMap>(this TPropertyMap propertyMap, ICellValueMapper mapper) where TPropertyMap : IMapper
    {
        var index = propertyMap.Mappers.Count;
        if (propertyMap.Mappers.Count > 0 && propertyMap.Mappers[propertyMap.Mappers.Count - 1] is InvalidCellMapper)
        {
            index--;
        }

        propertyMap.Mappers.Insert(index, mapper);
    }
}

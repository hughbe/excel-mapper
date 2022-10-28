using ExcelMapper.Abstractions;
using ExcelMapper.Fallbacks;
using ExcelMapper.Mappers;
using ExcelMapper.Readers;

namespace ExcelMapper;
 
public delegate T ConvertUsingSimpleMapperDelegate<out T>(object value);

/// <summary>
/// Extensions on IOneToOneMap to enable fluent "With" method chaining.
/// </summary>
public static class IOneToOneMapExtensions
{
    /// <summary>
    /// Makes the reader of the property map optional. For example, if the column doesn't exist
    /// or the index is invalid, an exception will not be thrown.
    /// </summary>
    /// <typeparam name="T">The type of the property map.</typeparam>
    /// <param name="propertyMap">The property map to use.</param>
    /// <returns>The property map on which this method was invoked.</returns>
    public static IOneToOneMap<T> MakeOptional<T>(this IOneToOneMap<T> propertyMap)
    {
        propertyMap.Optional = true;
        return propertyMap;
    }

    /// <summary>
    /// Sets the reader of the property map to read the value of a single cell contained in the column with
    /// the given names.
    /// </summary>
    /// <typeparam name="T">The type of the property map.</typeparam>
    /// <param name="propertyMap">The property map to use.</param>
    /// <param name="predicate">A predicate which returns whether a Column Name was matched or not</param>
    /// <returns>The property map on which this method was invoked.</returns>
    public static IOneToOneMap<T> WithColumnNameMatching<T>(this IOneToOneMap<T> propertyMap, Func<string, bool> predicate)
    {
        return propertyMap.WithReader(new ColumnNameMatchingValueReader(predicate));
    }

    /// <summary>
    /// Sets the reader of the property map to read the value of a single cell contained in the column with
    /// the given name.
    /// </summary>
    /// <typeparam name="T">The type of the property map.</typeparam>
    /// <param name="propertyMap">The property map to use.</param>
    /// <param name="columnName">The name of the column to read</param>
    /// <returns>The property map on which this method was invoked.</returns>
    public static IOneToOneMap<T> WithColumnName<T>(this IOneToOneMap<T> propertyMap, string columnName)
    {
        return propertyMap
            .WithReader(new ColumnNameValueReader(columnName));
    }

    /// <summary>
    /// Sets the reader of the property map to read the value of a single cell contained in the column at
    /// the given zero-based index.
    /// </summary>
    /// <typeparam name="T">The type of the property map.</typeparam>
    /// <param name="propertyMap">The property map to use.</param>
    /// <param name="columnIndex">The zero-based index of the column to read</param>
    /// <returns>The property map on which this method was invoked.</returns>
    public static IOneToOneMap<T> WithColumnIndex<T>(this IOneToOneMap<T> propertyMap, int columnIndex)
    {
        return propertyMap
            .WithReader(new ColumnIndexValueReader(columnIndex));
    }

    /// <summary>
    /// Sets the reader of the property map to use a custom cell value reader.
    /// </summary>
    /// <typeparam name="T">The type of the property map.</typeparam>
    /// <param name="propertyMap">The property map to use.</param>
    /// <param name="reader">The custom reader to use.</param>
    /// <returns>The property map on which this method was invoked.</returns>
    public static IOneToOneMap<T> WithReader<T>(this IOneToOneMap<T> propertyMap, ICellReader reader)
    {
        if (reader == null)
        {
            throw new ArgumentNullException(nameof(reader));
        }

        propertyMap.Reader = reader;
        return propertyMap;
    }
       
    /// <summary>
    /// Specifies that the string value of the cell should be trimmed before it is mapped to
    /// a property or field.
    /// </summary>
    /// <typeparam name="T">The type of the property map.</typeparam>
    /// <param name="propertyMap">The property map to use.</param>
    /// <returns>The property map on which this method was invoked.</returns>
    public static IOneToOneMap<T> WithTrim<T>(this IOneToOneMap<T> propertyMap)
    {
        propertyMap.Mappers.Insert(0, new TrimCellValueTransformer());
        return propertyMap;
    }

    /// <summary>
    /// Specifies additional custom mappers that will be used to map the value of a cell to
    /// a property or field.
    /// </summary>
    /// <typeparam name="T">The type of the property map.</typeparam>
    /// <param name="propertyMap">The property map to use.</param>
    /// <param name="mappers">A list of additional custom mappers that will be used to map the value of a cell to a property or field</param>
    /// <returns>The property map on which this method was invoked.</returns>
    public static IOneToOneMap<T> WithCellValueMappers<T>(this IOneToOneMap<T> propertyMap, params ICellValueMapper[] mappers)
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
    /// <typeparam name="T">The type of the property map.</typeparam>
    /// <typeparam name="T">The type of the property or field that the property map represents.</typeparam>
    /// <param name="propertyMap">The property map to use.</param>
    /// <param name="mappingDictionary">A dictionary that maps a fixed string value to a fixed value of T.</param>
    /// <param name="comparer">The comparer uses to map fixed string values. This allows for case-insensitive mappings, for example.</param>
    /// <returns>The property map on which this method was invoked.</returns>
    public static IOneToOneMap<T> WithMapping<T>(this IOneToOneMap<T> propertyMap, IDictionary<string, T> mappingDictionary, IEqualityComparer<string> comparer = null)
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
    public static IOneToOneMap<DateTime> WithDateFormats(this IOneToOneMap<DateTime> propertyMap, params string[] formats)
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
    public static IOneToOneMap<DateTime> WithDateFormats(this IOneToOneMap<DateTime> propertyMap, IEnumerable<string> formats)
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
    public static IOneToOneMap<DateTime?> WithDateFormats(this IOneToOneMap<DateTime?> propertyMap, params string[] formats)
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
    public static IOneToOneMap<DateTime?> WithDateFormats(this IOneToOneMap<DateTime?> propertyMap, IEnumerable<string> formats)
    {
        return propertyMap.WithDateFormats(formats?.ToArray());
    }

    private static void AddFormats<T>(this IOneToOneMap<T> propertyMap, string[] formats)
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
    /// <typeparam name="T">The type of the property map.</typeparam>
    /// <typeparam name="T">The type of the property or field that the property map represents.</typeparam>
    /// <param name="propertyMap">The property map to use.</param>
    /// <param name="converter">A delegate that is invoked to map the string value of a cell to the value of a property or field.</param>
    /// <returns>The property map on which this method was invoked.</returns>
    public static IOneToOneMap<T> WithConverter<T>(this IOneToOneMap<T> propertyMap, ConvertUsingSimpleMapperDelegate<T> converter)
    {
        if (converter == null)
        {
            throw new ArgumentNullException(nameof(converter));
        }

        ConvertUsingMapperDelegate actualConverter = (ExcelCell cell, CellValueMapperResult previous) =>
        {
            try
            {
                object result = converter(previous.Value);
                return previous.Success(result);
            }
            catch (Exception exception)
            {
                return previous.Invalid(exception);
            }
        };

        return propertyMap.WithConverter(actualConverter);
    }

    public static IOneToOneMap<T> WithConverter<T>(this IOneToOneMap<T> propertyMap, ConvertUsingMapperDelegate converter)
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
    /// <typeparam name="T">The type of the property map.</typeparam>
    /// <param name="propertyMap">The property map to use.</param>
    /// <param name="defaultValue">The value that will be assigned to the property or field if the value of a cell is empty or cannot be mapped.</param>
    /// <returns>The property map on which this method was invoked.</returns>
    public static IOneToOneMap<T> WithValueFallback<T>(this IOneToOneMap<T> propertyMap, object defaultValue)
    {
        return propertyMap
            .WithEmptyFallback(defaultValue)
            .WithInvalidFallback(defaultValue);
    }

    /// <summary>
    /// Specifies that the property map should throw an exception if the value of a cell if empty or cannot be mapped.
    /// </summary>
    /// <typeparam name="T">The type of the property map.</typeparam>
    /// <param name="propertyMap">The property map to use.</param>
    /// <returns>The property map on which this method was invoked.</returns>
    public static IOneToOneMap<T> WithThrowingFallback<T>(this IOneToOneMap<T> propertyMap)
    {
        return propertyMap
            .WithThrowingEmptyFallback()
            .WithThrowingInvalidFallback();
    }

    /// <summary>
    /// Specifies a fixed fallback to be used if the value of a cell is empty.
    /// </summary>
    /// <typeparam name="T">The type of the property map.</typeparam>
    /// <param name="propertyMap">The property map to use.</param>
    /// <param name="fallbackValue">The value that will be assigned to the property or field if the value of a cell is empty.</param>
    /// <returns>The property map on which this method was invoked.</returns>
    public static IOneToOneMap<T> WithEmptyFallback<T>(this IOneToOneMap<T> propertyMap, object fallbackValue)
    {
        return propertyMap
            .WithEmptyFallbackItem(new FixedValueFallback(fallbackValue));
    }

    /// <summary>
    /// Specifies a custom fallback to be used if the value of a cell is empty.
    /// </summary>
    /// <typeparam name="T">The type of the property map.</typeparam>
    /// <param name="propertyMap">The property map to use.</param>
    /// <param name="fallbackItem">The fallback to be used if the value of a cell is empty.</param>
    /// <returns>The property map on which this method was invoked.</returns>
    public static IOneToOneMap<T> WithEmptyFallbackItem<T>(this IOneToOneMap<T> propertyMap, IEmptyCellFallback fallbackItem)
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
    /// <typeparam name="T">The type of the property map.</typeparam>
    /// <param name="propertyMap">The property map to use.</param>
    /// <returns>The property map on which this method was invoked.</returns>
    public static IOneToOneMap<T> WithThrowingEmptyFallback<T>(this IOneToOneMap<T> propertyMap)
    {
        return propertyMap
            .WithEmptyFallbackItem(new ThrowFallback());
    }

    /// <summary>
    /// Specifies that the property map should throw an exception if the value of a cell cannot be mapped.
    /// </summary>
    /// <typeparam name="T">The type of the property map.</typeparam>
    /// <param name="propertyMap">The property map to use.</param>
    /// <returns>The property map on which this method was invoked.</returns>
    public static IOneToOneMap<T> WithThrowingInvalidFallback<T>(this IOneToOneMap<T> propertyMap)
    {
        return propertyMap
            .WithInvalidFallbackItem(new ThrowFallback());
    }

    /// <summary>
    /// Specifies a fixed fallback to be used if the value of a cell cannot be mapped.
    /// </summary>
    /// <typeparam name="T">The type of the property map.</typeparam>
    /// <param name="propertyMap">The property map to use.</param>
    /// <param name="fallbackValue">The value that will be assigned to the property or field if the value of a cell cannot be mapped.</param>
    /// <returns>The property map on which this method was invoked.</returns>
    public static IOneToOneMap<T> WithInvalidFallback<T>(this IOneToOneMap<T> propertyMap, object fallbackValue)
    {
        return propertyMap
            .WithInvalidFallbackItem(new FixedValueFallback(fallbackValue));
    }

    /// <summary>
    /// Specifies a custom fallback to be used if the value of a cell cannot be mapped.
    /// </summary>
    /// <typeparam name="T">The type of the property map.</typeparam>
    /// <param name="propertyMap">The property map to use.</param>
    /// <param name="fallbackItem">The fallback to be used if the value of a cell cannot be mapped.</param>
    /// <returns>The property map on which this method was invoked.</returns>
    public static IOneToOneMap<T> WithInvalidFallbackItem<T>(this IOneToOneMap<T> propertyMap, IInvalidCellFallback fallbackItem)
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

    private static void AppendMapper<T>(this IOneToOneMap<T> propertyMap, ICellValueMapper mapper)
    {
        var index = propertyMap.Mappers.Count;
        if (propertyMap.Mappers.Count > 0 && propertyMap.Mappers[propertyMap.Mappers.Count - 1] is InvalidCellMapper)
        {
            index--;
        }

        propertyMap.Mappers.Insert(index, mapper);
    }
}

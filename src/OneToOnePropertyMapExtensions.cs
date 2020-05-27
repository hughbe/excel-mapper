using System;
using System.Collections.Generic;
using System.Linq;
using ExcelMapper.Mappings;
using ExcelMapper.Mappings.Fallbacks;
using ExcelMapper.Mappings.Mappers;
using ExcelMapper.Mappings.Readers;
using ExcelMapper.Mappings.Support;
using ExcelMapper.Mappings.Transformers;

namespace ExcelMapper
{
    public delegate T ConvertUsingSimpleMapperDelegate<out T>(string stringValue);

    /// <summary>
    /// Extensions on OneToOnePropertyMap to enable fluent "With" method chaining.
    /// </summary>
    public static class OneToOnePropertyMapExtensions
    {
        /// <summary>
        /// Sets the reader of the property map to read the value of a single cell contained in the column with
        /// the given names.
        /// </summary>
        /// <typeparam name="T">The type of the property map.</typeparam>
        /// <param name="propertyMap">The property map to use.</param>
        /// <param name="predicate">A predicate which returns whether a Column Name was matched or not</param>
        /// <returns>The property map on which this method was invoked.</returns>
        public static T WithColumnNameMatching<T>(this T propertyMap, Func<string, bool> predicate)
            where T : IOneToOnePropertyMap
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
        public static T WithColumnName<T>(this T propertyMap, string columnName) where T : IOneToOnePropertyMap
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
        public static T WithColumnIndex<T>(this T propertyMap, int columnIndex) where T : IOneToOnePropertyMap
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
        public static T WithReader<T>(this T propertyMap, ISingleCellValueReader reader) where T : IOneToOnePropertyMap
        {
            if (reader == null)
            {
                throw new ArgumentNullException(nameof(reader));
            }

            propertyMap.CellReader = reader;
            return propertyMap;
        }

        /// <summary>
        /// Makes the reader of the property map optional. For example, if the column doesn't exist
        /// or the index is invalid, an exception will not be thrown.
        /// </summary>
        /// <typeparam name="T">The type of the property map.</typeparam>
        /// <param name="propertyMap">The property map to use.</param>
        /// <returns>The property map on which this method was invoked.</returns>
        public static T MakeOptional<T>(this T propertyMap) where T : IOneToOnePropertyMap
        {
            if (propertyMap.Optional)
            {
                return propertyMap;
            }

            propertyMap.Optional = true;
            return propertyMap;
        }

        /// <summary>
        /// Specifies that the string value of the cell should be trimmed before it is mapped to
        /// a property or field.
        /// </summary>
        /// <typeparam name="T">The type of the property map.</typeparam>
        /// <param name="propertyMap">The property map to use.</param>
        /// <returns>The property map on which this method was invoked.</returns>
        public static T WithTrim<T>(this T propertyMap) where T : IOneToOnePropertyMap
        {
            var transformer = new TrimCellValueTransformer();
            propertyMap.AddCellValueTransformer(transformer);
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
        public static TPropertyMap WithCellValueMappers<TPropertyMap>(this TPropertyMap propertyMap, params ICellValueMapper[] mappers) where TPropertyMap : IOneToOnePropertyMap
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
                propertyMap.AddCellValueMapper(mapper);
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
        public static TPropertyMap WithMapping<TPropertyMap, T>(this TPropertyMap propertyMap, IDictionary<string, T> mappingDictionary, IEqualityComparer<string> comparer = null) where TPropertyMap : IOneToOnePropertyMap<T>
        {
            var item = new DictionaryMapper<T>(mappingDictionary, comparer);
            propertyMap.AddCellValueMapper(item);
            return propertyMap;
        }

        /// <summary>
        /// Specifies data formats used when mapping the value of a cell to a DateTime. This is useful for
        /// mapping columns where data formats differ. Existing date formats are overriden.
        /// </summary>
        /// <param name="propertyMap">The property map to use.</param>
        /// <param name="formats">A list of date formats to use when mapping the value of a cell to a DateTime.</param>
        /// <returns>The property map on which this method was invoked.</returns>
        public static OneToOnePropertyMap<DateTime> WithDateFormats(this OneToOnePropertyMap<DateTime> propertyMap, params string[] formats)
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
        public static OneToOnePropertyMap<DateTime> WithDateFormats(this OneToOnePropertyMap<DateTime> propertyMap, IEnumerable<string> formats)
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
        public static OneToOnePropertyMap<DateTime?> WithDateFormats(this OneToOnePropertyMap<DateTime?> propertyMap, params string[] formats)
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
        public static OneToOnePropertyMap<DateTime?> WithDateFormats(this OneToOnePropertyMap<DateTime?> propertyMap, IEnumerable<string> formats)
        {
            return propertyMap.WithDateFormats(formats?.ToArray());
        }

        private static void AddFormats(this IOneToOnePropertyMap propertyMap, string[] formats)
        {
            if (formats == null)
            {
                throw new ArgumentNullException(nameof(formats));
            }

            if (formats.Length == 0)
            {
                throw new ArgumentException("Formats cannot be empty.", nameof(formats));
            }

            DateTimeMapper dateTimeItem = (DateTimeMapper)propertyMap.CellValueMappers.FirstOrDefault(item => item is DateTimeMapper);
            if (dateTimeItem == null)
            {
                dateTimeItem = new DateTimeMapper();
                propertyMap.AddCellValueMapper(dateTimeItem);
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
        public static TPropertyMap WithConverter<TPropertyMap, T>(this TPropertyMap propertyMap, ConvertUsingSimpleMapperDelegate<T> converter) where TPropertyMap : IOneToOnePropertyMap<T>
        {
            if (converter == null)
            {
                throw new ArgumentNullException(nameof(converter));
            }

            ConvertUsingMapperDelegate actualConverter = (ReadCellValueResult mapResult, ref object value) =>
            {
                try
                {
                    value = converter(mapResult.StringValue);
                    return PropertyMapperResultType.Success;
                }
                catch
                {
                    return PropertyMapperResultType.Invalid;
                }
            };

            var item = new ConvertUsingMapper(actualConverter);
            propertyMap.AddCellValueMapper(item);
            return propertyMap;
        }

        /// <summary>
        /// Specifies a fixed fallback to be used if the value of a cell is empty or cannot be mapped.
        /// </summary>
        /// <typeparam name="TPropertyMap">The type of the property map.</typeparam>
        /// <param name="propertyMap">The property map to use.</param>
        /// <param name="defaultValue">The value that will be assigned to the property or field if the value of a cell is empty or cannot be mapped.</param>
        /// <returns>The property map on which this method was invoked.</returns>
        public static TPropertyMap WithValueFallback<TPropertyMap>(this TPropertyMap propertyMap, object defaultValue) where TPropertyMap : IOneToOnePropertyMap
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
        public static TPropertyMap WithThrowingFallback<TPropertyMap>(this TPropertyMap propertyMap) where TPropertyMap : IOneToOnePropertyMap
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
        public static TPropertyMap WithEmptyFallback<TPropertyMap>(this TPropertyMap propertyMap, object fallbackValue) where TPropertyMap : IOneToOnePropertyMap
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
        public static TPropertyMap WithEmptyFallbackItem<TPropertyMap>(this TPropertyMap propertyMap, IFallbackItem fallbackItem) where TPropertyMap : IOneToOnePropertyMap
        {
            propertyMap.EmptyFallback = fallbackItem ?? throw new ArgumentNullException(nameof(fallbackItem));
            return propertyMap;
        }

        /// <summary>
        /// Specifies that the property map should throw an exception if the value of a cell is empty.
        /// </summary>
        /// <typeparam name="TPropertyMap">The type of the property map.</typeparam>
        /// <param name="propertyMap">The property map to use.</param>
        /// <returns>The property map on which this method was invoked.</returns>
        public static TPropertyMap WithThrowingEmptyFallback<TPropertyMap>(this TPropertyMap propertyMap) where TPropertyMap : IOneToOnePropertyMap
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
        public static TPropertyMap WithThrowingInvalidFallback<TPropertyMap>(this TPropertyMap propertyMap) where TPropertyMap : IOneToOnePropertyMap
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
        public static TPropertyMap WithInvalidFallback<TPropertyMap>(this TPropertyMap propertyMap, object fallbackValue) where TPropertyMap : IOneToOnePropertyMap
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
        public static TPropertyMap WithInvalidFallbackItem<TPropertyMap>(this TPropertyMap propertyMap, IFallbackItem fallbackItem) where TPropertyMap : IOneToOnePropertyMap
        {
            propertyMap.InvalidFallback = fallbackItem ?? throw new ArgumentNullException(nameof(fallbackItem));
            return propertyMap;
        }
    }
}

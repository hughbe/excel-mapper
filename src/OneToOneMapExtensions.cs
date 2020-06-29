using System;
using ExcelMapper.Abstractions;
using ExcelMapper.Readers;

namespace ExcelMapper
{
    /// <summary>
    /// Extensions on OneToOneMap to enable fluent "With" method chaining.
    /// </summary>
    public static class OneToOneMapExtensions
    {
        /// <summary>
        /// Makes the reader of the property map optional. For example, if the column doesn't exist
        /// or the index is invalid, an exception will not be thrown.
        /// </summary>
        /// <typeparam name="T">The type of the property map.</typeparam>
        /// <param name="propertyMap">The property map to use.</param>
        /// <returns>The property map on which this method was invoked.</returns>
        public static OneToOneMap<T> MakeOptional<T>(this OneToOneMap<T> propertyMap)
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
        public static OneToOneMap<T> WithColumnNameMatching<T>(this OneToOneMap<T> propertyMap, Func<string, bool> predicate)
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
        public static OneToOneMap<T> WithColumnName<T>(this OneToOneMap<T> propertyMap, string columnName)
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
        public static OneToOneMap<T> WithColumnIndex<T>(this OneToOneMap<T> propertyMap, int columnIndex)
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
        public static OneToOneMap<T> WithReader<T>(this OneToOneMap<T> propertyMap, ISingleCellValueReader reader)
        {
            if (reader == null)
            {
                throw new ArgumentNullException(nameof(reader));
            }

            propertyMap.CellReader = reader;
            return propertyMap;
        }
    }
}

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
        public static T MakeOptional<T>(this T propertyMap) where T : OneToOneMap
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
        public static T WithColumnNameMatching<T>(this T propertyMap, Func<string, bool> predicate)
            where T : OneToOneMap
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
        public static T WithColumnName<T>(this T propertyMap, string columnName) where T : OneToOneMap
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
        public static T WithColumnIndex<T>(this T propertyMap, int columnIndex) where T : OneToOneMap
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
        public static T WithReader<T>(this T propertyMap, ISingleCellValueReader reader) where T : OneToOneMap
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

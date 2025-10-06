using System;
using System.Collections.Generic;
using ExcelMapper.Abstractions;
using ExcelMapper.Readers;

namespace ExcelMapper;

public static class IManyToOneMapExtensions
{
    /// <summary>
    /// Sets the reader of the map to read the values of one or more cells contained
    /// in the columns with the given names.
    /// </summary>
    /// <param name="columnNames">The name of each column to read.</param>
    /// <returns>The map that invoked this method.</returns>
    public static TPropertyMap WithColumnNames<TPropertyMap>(this TPropertyMap map, params string[] columnNames) where TPropertyMap : IManyToOneMap
        => map.WithReaderFactory(new ColumnNamesReaderFactory(columnNames));

    /// <summary>
    /// Sets the reader of the map to read the values of one or more cells contained
    /// in the columns with the given names.
    /// </summary>
    /// <param name="columnNames">The name of each column to read.</param>
    /// <returns>The map that invoked this method.</returns>
    public static TPropertyMap WithColumnNames<TPropertyMap>(this TPropertyMap map, IEnumerable<string> columnNames) where TPropertyMap : IManyToOneMap
    {
        if (columnNames == null)
        {
            throw new ArgumentNullException(nameof(columnNames));
        }

        return map.WithColumnNames([.. columnNames]);
    }

    /// <summary>
    /// Sets the reader of the map to read the values of one or more cells contained
    /// in the columns matching the result of IExcelColumnMatcher.ColumnMatches.
    /// </summary>
    /// <param name="matcher">The matcher of each column to read.</param>
    /// <returns>The map that invoked this method.</returns>
    public static TPropertyMap WithColumnsMatching<TPropertyMap>(this TPropertyMap map, IExcelColumnMatcher matcher) where TPropertyMap : IManyToOneMap
        => map.WithReaderFactory(new ColumnsMatchingReaderFactory(matcher));

    /// <summary>
    /// Sets the reader of the map to read the values of one or more cells contained
    /// in the columns with the given zero-based indices.
    /// </summary>
    /// <param name="columnIndices">The zero-based index of each column to read.</param>
    /// <returns>The map that invoked this method.</returns>
    public static TPropertyMap WithColumnIndices<TPropertyMap>(this TPropertyMap map, params int[] columnIndices) where TPropertyMap : IManyToOneMap
        => map.WithReaderFactory(new ColumnIndicesReaderFactory(columnIndices));

    /// <summary>
    /// Sets the reader of the map to read the values of one or more cells contained
    /// in the columns with the given zero-based indices.
    /// </summary>
    /// <param name="columnIndices">The zero-based index of each column to read.</param>
    /// <returns>The map that invoked this method.</returns>
    public static TPropertyMap WithColumnIndices<TPropertyMap>(this TPropertyMap map, IEnumerable<int> columnIndices) where TPropertyMap : IManyToOneMap
    {
        if (columnIndices == null)
        {
            throw new ArgumentNullException(nameof(columnIndices));
        }
        
        return map.WithColumnIndices([.. columnIndices]);
    }
    /// <summary>
    /// Sets the reader of the map to use a custom cell values reader.
    /// </summary>
    /// <typeparam name="TPropertyMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="reader">The custom reader to use.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TPropertyMap WithReaderFactory<TPropertyMap>(this TPropertyMap map, ICellsReaderFactory readerFactory) where TPropertyMap : IManyToOneMap
    {
        map.ReaderFactory = readerFactory ?? throw new ArgumentNullException(nameof(readerFactory));
        return map;
    }
}

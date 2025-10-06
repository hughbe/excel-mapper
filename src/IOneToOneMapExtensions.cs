using System;
using ExcelMapper.Abstractions;
using ExcelMapper.Readers;

namespace ExcelMapper;

public static class IOneToOneMapExtensions
{
    /// <summary>
    /// Sets the reader of the map to read the value of a single cell with a column name
    /// matching a predicate.
    /// </summary>
    /// <typeparam name="TPropertyMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="predicate">A predicate which returns whether a Column Name was matched or not</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TPropertyMap WithColumnNameMatching<TPropertyMap>(this TPropertyMap map, Func<string, bool> predicate) where TPropertyMap : IOneToOneMap
        => map.WithColumnMatching(new PredicateColumnMatcher(predicate));

    /// <summary>
    /// Sets the reader of the map to read the value of a single cell for the first column matching
    /// IExcelColumnMatcher.ColumnMatches(ExcelSheet sheet, int columnIndex).
    /// </summary>
    /// <typeparam name="TPropertyMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="matcher">An IExcelColumnMatcher that returns whether a Column Name was matched or not</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TPropertyMap WithColumnMatching<TPropertyMap>(this TPropertyMap map, IExcelColumnMatcher matcher) where TPropertyMap : IOneToOneMap
        => map.WithReaderFactory(new ColumnsMatchingReaderFactory(matcher));

    /// <summary>
    /// Sets the reader of the map to read the value of a single cell contained in the column with
    /// the given names.
    /// </summary>
    /// <typeparam name="TPropertyMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="predicate">A predicate which returns whether a Column Name was matched or not</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TPropertyMap WithColumnNameMatching<TPropertyMap>(this TPropertyMap map, params string[] columnNames) where TPropertyMap : IOneToOneMap
        => map.WithReaderFactory(new ColumnNamesReaderFactory(columnNames));

    /// <summary>
    /// Sets the reader of the map to read the value of a single cell contained in the column with
    /// the given name.
    /// </summary>
    /// <typeparam name="TPropertyMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="columnName">The name of the column to read</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TPropertyMap WithColumnName<TPropertyMap>(this TPropertyMap map, string columnName) where TPropertyMap : IOneToOneMap
        => map.WithReaderFactory(new ColumnNameReaderFactory(columnName));

    /// <summary>
    /// Sets the reader of the map to read the value of a single cell contained in the column at
    /// the given zero-based index.
    /// </summary>
    /// <typeparam name="TPropertyMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="columnIndex">The zero-based index of the column to read</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TPropertyMap WithColumnIndex<TPropertyMap>(this TPropertyMap map, int columnIndex) where TPropertyMap : IOneToOneMap
        => map.WithReaderFactory(new ColumnIndexReaderFactory(columnIndex));

    /// <summary>
    /// Sets the reader of the map to use a custom cell value reader.
    /// </summary>
    /// <typeparam name="TPropertyMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="reader">The custom reader to use.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TPropertyMap WithReaderFactory<TPropertyMap>(this TPropertyMap map, ICellReaderFactory readerFactory) where TPropertyMap : IOneToOneMap
    {
        map.ReaderFactory = readerFactory ?? throw new ArgumentNullException(nameof(readerFactory));
        return map;
    }
}
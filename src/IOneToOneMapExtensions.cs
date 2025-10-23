using ExcelMapper.Readers;

namespace ExcelMapper;

/// <summary>
/// Extension methods for <see cref="IOneToOneMap"/>.
/// </summary>
public static class IOneToOneMapExtensions
{
    /// <summary>
    /// Sets the reader of the map to read the value of a single cell with a column name
    /// matching a predicate.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="predicate">A predicate which returns whether a Column Name was matched or not</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithColumnNameMatching<TMap>(this TMap map, Func<string, bool> predicate) where TMap : IOneToOneMap
        => map.WithColumnMatching(new PredicateColumnMatcher(predicate));

    /// <summary>
    /// Sets the reader of the map to read the value of a single cell for the first column matching
    /// IExcelColumnMatcher.ColumnMatches(ExcelSheet sheet, int columnIndex).
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="matcher">An IExcelColumnMatcher that returns whether a Column Name was matched or not</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithColumnMatching<TMap>(this TMap map, IExcelColumnMatcher matcher) where TMap : IOneToOneMap
        => map.WithReaderFactory(new ColumnsMatchingReaderFactory(matcher));

    /// <summary>
    /// Sets the reader of the map to read the value of a single cell contained in the column with
    /// the given names.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="predicate">A predicate which returns whether a Column Name was matched or not</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithColumnNameMatching<TMap>(this TMap map, params string[] columnNames) where TMap : IOneToOneMap
        => map.WithReaderFactory(new ColumnNamesReaderFactory(columnNames));

    /// <summary>
    /// Sets the reader of the map to read the value of a single cell contained in the column with
    /// the given name.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="columnName">The name of the column to read</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithColumnName<TMap>(this TMap map, string columnName) where TMap : IOneToOneMap
        => map.WithReaderFactory(new ColumnNameReaderFactory(columnName));

    /// <summary>
    /// Sets the reader of the map to read the value of a single cell contained in the column at
    /// the given zero-based index.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="columnIndex">The zero-based index of the column to read</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithColumnIndex<TMap>(this TMap map, int columnIndex) where TMap : IOneToOneMap
        => map.WithReaderFactory(new ColumnIndexReaderFactory(columnIndex));

    /// <summary>
    /// Sets the reader of the map to use a custom cell value reader.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="reader">The custom reader to use.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithReaderFactory<TMap>(this TMap map, ICellReaderFactory readerFactory) where TMap : IOneToOneMap
    {
        ArgumentNullException.ThrowIfNull(readerFactory);
        map.ReaderFactory = readerFactory;
        return map;
    }
}
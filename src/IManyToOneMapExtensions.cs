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
    public static TMap WithColumnNames<TMap>(this TMap map, params string[] columnNames) where TMap : IManyToOneMap
        => map.WithReaderFactory(new ColumnNamesReaderFactory(columnNames));

    /// <summary>
    /// Sets the reader of the map to read the values of one or more cells contained
    /// in the columns with the given names.
    /// </summary>
    /// <param name="columnNames">The name of each column to read.</param>
    /// <returns>The map that invoked this method.</returns>
    public static TMap WithColumnNames<TMap>(this TMap map, IEnumerable<string> columnNames) where TMap : IManyToOneMap
    {
        ArgumentNullException.ThrowIfNull(columnNames);

        return map.WithColumnNames([.. columnNames]);
    }

    /// <summary>
    /// Sets the reader of the map to read the values of one or more cells contained
    /// in the columns matching the result of IExcelColumnMatcher.ColumnMatches.
    /// </summary>
    /// <param name="matcher">The matcher of each column to read.</param>
    /// <returns>The map that invoked this method.</returns>
    public static TMap WithColumnsMatching<TMap>(this TMap map, IExcelColumnMatcher matcher) where TMap : IManyToOneMap
        => map.WithReaderFactory(new ColumnsMatchingReaderFactory(matcher));

    /// <summary>
    /// Sets the reader of the map to read the values of one or more cells contained
    /// in the columns with the given zero-based indices.
    /// </summary>
    /// <param name="columnIndices">The zero-based index of each column to read.</param>
    /// <returns>The map that invoked this method.</returns>
    public static TMap WithColumnIndices<TMap>(this TMap map, params int[] columnIndices) where TMap : IManyToOneMap
        => map.WithReaderFactory(new ColumnIndicesReaderFactory(columnIndices));

    /// <summary>
    /// Sets the reader of the map to read the values of one or more cells contained
    /// in the columns with the given zero-based indices.
    /// </summary>
    /// <param name="columnIndices">The zero-based index of each column to read.</param>
    /// <returns>The map that invoked this method.</returns>
    public static TMap WithColumnIndices<TMap>(this TMap map, IEnumerable<int> columnIndices) where TMap : IManyToOneMap
    {
        ArgumentNullException.ThrowIfNull(columnIndices);

        return map.WithColumnIndices([.. columnIndices]);
    }
    /// <summary>
    /// Sets the reader of the map to use a custom cell values reader.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="reader">The custom reader to use.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithReaderFactory<TMap>(this TMap map, ICellsReaderFactory readerFactory) where TMap : IManyToOneMap
    {
        ArgumentNullException.ThrowIfNull(readerFactory);
        map.ReaderFactory = readerFactory;
        return map;
    }
}

namespace ExcelMapper;

public static class IToOneMapExtensions
{
    /// <summary>
    /// Makes the reader of the map optional. For example, if the column doesn't exist
    /// or the index is invalid, an exception will not be thrown.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap MakeOptional<TMap>(this TMap map) where TMap : IToOneMap
    {
        map.Optional = true;
        return map;
    }
    
    /// <summary>
    /// Makes the reader of the map peserve formatting when reading string values.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap MakePreserveFormatting<TMap>(this TMap map) where TMap : IToOneMap
    {
        map.PreserveFormatting = true;
        return map;
    }
}
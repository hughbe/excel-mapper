using System.Globalization;
using ExcelMapper.Mappers;

namespace ExcelMapper;

// The number options have an overload for each built-in number type and its nullable version, in
// every build: the netstandard builds don't have INumberBase<T>, and a library built against them
// must find the same methods on .NET 8, where these are chosen over the generic overloads in
// ValuePipelineExtensions.cs.
public static partial class ValuePipelineExtensions
{
#if NET7_0_OR_GREATER
    private static void AddNumberMapperIfMissing<T>(CellPipeline<T> pipeline) where T : struct, System.Numerics.INumberBase<T>
        => pipeline.AddMapperIfMissing<T, StructNumberMapper<T>>(static () => new StructNumberMapper<T>());

    private static void AddNumberMapperIfMissing<T>(CellPipeline<T?> pipeline) where T : struct, System.Numerics.INumberBase<T>
        => pipeline.AddMapperIfMissing<T?, StructNumberMapper<T>>(static () => new StructNumberMapper<T>());
#else
    private static void AddNumberMapperIfMissing<T>(CellPipeline<T> pipeline) where T : struct
        => pipeline.AddMapperIfMissing<T, NumberMapper<T>>(static () => new NumberMapper<T>());

    private static void AddNumberMapperIfMissing<T>(CellPipeline<T?> pipeline) where T : struct
        => pipeline.AddMapperIfMissing<T?, NumberMapper<T>>(static () => new NumberMapper<T>());
#endif

    /// <summary>
    /// Specifies the number styles to use when mapping the value of a cell to a <see cref="sbyte"/>.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="style">The number styles to use when parsing a cell's text.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithNumberStyle<TMap>(this ValuePipeline<sbyte, TMap> map, NumberStyles style) where TMap : ValuePipeline<sbyte, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.NumberStyle = style;
        AddNumberMapperIfMissing(map.Pipeline);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the format provider to use when mapping the value of a cell to a <see cref="sbyte"/>.
    /// This is useful for mapping columns where cultures differ.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider to use when parsing a cell's text.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormatProvider<TMap>(this ValuePipeline<sbyte, TMap> map, IFormatProvider? provider) where TMap : ValuePipeline<sbyte, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.Provider = provider;
        AddNumberMapperIfMissing(map.Pipeline);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the number styles to use when mapping the value of a cell to a <see cref="sbyte"/> or null.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="style">The number styles to use when parsing a cell's text.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithNumberStyle<TMap>(this ValuePipeline<sbyte?, TMap> map, NumberStyles style) where TMap : ValuePipeline<sbyte?, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.NumberStyle = style;
        AddNumberMapperIfMissing(map.Pipeline);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the format provider to use when mapping the value of a cell to a <see cref="sbyte"/> or null.
    /// This is useful for mapping columns where cultures differ.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider to use when parsing a cell's text.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormatProvider<TMap>(this ValuePipeline<sbyte?, TMap> map, IFormatProvider? provider) where TMap : ValuePipeline<sbyte?, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.Provider = provider;
        AddNumberMapperIfMissing(map.Pipeline);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the number styles to use when mapping the value of a cell to a <see cref="byte"/>.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="style">The number styles to use when parsing a cell's text.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithNumberStyle<TMap>(this ValuePipeline<byte, TMap> map, NumberStyles style) where TMap : ValuePipeline<byte, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.NumberStyle = style;
        AddNumberMapperIfMissing(map.Pipeline);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the format provider to use when mapping the value of a cell to a <see cref="byte"/>.
    /// This is useful for mapping columns where cultures differ.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider to use when parsing a cell's text.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormatProvider<TMap>(this ValuePipeline<byte, TMap> map, IFormatProvider? provider) where TMap : ValuePipeline<byte, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.Provider = provider;
        AddNumberMapperIfMissing(map.Pipeline);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the number styles to use when mapping the value of a cell to a <see cref="byte"/> or null.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="style">The number styles to use when parsing a cell's text.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithNumberStyle<TMap>(this ValuePipeline<byte?, TMap> map, NumberStyles style) where TMap : ValuePipeline<byte?, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.NumberStyle = style;
        AddNumberMapperIfMissing(map.Pipeline);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the format provider to use when mapping the value of a cell to a <see cref="byte"/> or null.
    /// This is useful for mapping columns where cultures differ.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider to use when parsing a cell's text.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormatProvider<TMap>(this ValuePipeline<byte?, TMap> map, IFormatProvider? provider) where TMap : ValuePipeline<byte?, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.Provider = provider;
        AddNumberMapperIfMissing(map.Pipeline);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the number styles to use when mapping the value of a cell to a <see cref="short"/>.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="style">The number styles to use when parsing a cell's text.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithNumberStyle<TMap>(this ValuePipeline<short, TMap> map, NumberStyles style) where TMap : ValuePipeline<short, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.NumberStyle = style;
        AddNumberMapperIfMissing(map.Pipeline);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the format provider to use when mapping the value of a cell to a <see cref="short"/>.
    /// This is useful for mapping columns where cultures differ.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider to use when parsing a cell's text.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormatProvider<TMap>(this ValuePipeline<short, TMap> map, IFormatProvider? provider) where TMap : ValuePipeline<short, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.Provider = provider;
        AddNumberMapperIfMissing(map.Pipeline);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the number styles to use when mapping the value of a cell to a <see cref="short"/> or null.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="style">The number styles to use when parsing a cell's text.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithNumberStyle<TMap>(this ValuePipeline<short?, TMap> map, NumberStyles style) where TMap : ValuePipeline<short?, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.NumberStyle = style;
        AddNumberMapperIfMissing(map.Pipeline);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the format provider to use when mapping the value of a cell to a <see cref="short"/> or null.
    /// This is useful for mapping columns where cultures differ.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider to use when parsing a cell's text.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormatProvider<TMap>(this ValuePipeline<short?, TMap> map, IFormatProvider? provider) where TMap : ValuePipeline<short?, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.Provider = provider;
        AddNumberMapperIfMissing(map.Pipeline);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the number styles to use when mapping the value of a cell to a <see cref="ushort"/>.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="style">The number styles to use when parsing a cell's text.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithNumberStyle<TMap>(this ValuePipeline<ushort, TMap> map, NumberStyles style) where TMap : ValuePipeline<ushort, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.NumberStyle = style;
        AddNumberMapperIfMissing(map.Pipeline);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the format provider to use when mapping the value of a cell to a <see cref="ushort"/>.
    /// This is useful for mapping columns where cultures differ.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider to use when parsing a cell's text.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormatProvider<TMap>(this ValuePipeline<ushort, TMap> map, IFormatProvider? provider) where TMap : ValuePipeline<ushort, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.Provider = provider;
        AddNumberMapperIfMissing(map.Pipeline);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the number styles to use when mapping the value of a cell to a <see cref="ushort"/> or null.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="style">The number styles to use when parsing a cell's text.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithNumberStyle<TMap>(this ValuePipeline<ushort?, TMap> map, NumberStyles style) where TMap : ValuePipeline<ushort?, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.NumberStyle = style;
        AddNumberMapperIfMissing(map.Pipeline);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the format provider to use when mapping the value of a cell to a <see cref="ushort"/> or null.
    /// This is useful for mapping columns where cultures differ.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider to use when parsing a cell's text.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormatProvider<TMap>(this ValuePipeline<ushort?, TMap> map, IFormatProvider? provider) where TMap : ValuePipeline<ushort?, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.Provider = provider;
        AddNumberMapperIfMissing(map.Pipeline);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the number styles to use when mapping the value of a cell to a <see cref="int"/>.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="style">The number styles to use when parsing a cell's text.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithNumberStyle<TMap>(this ValuePipeline<int, TMap> map, NumberStyles style) where TMap : ValuePipeline<int, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.NumberStyle = style;
        AddNumberMapperIfMissing(map.Pipeline);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the format provider to use when mapping the value of a cell to a <see cref="int"/>.
    /// This is useful for mapping columns where cultures differ.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider to use when parsing a cell's text.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormatProvider<TMap>(this ValuePipeline<int, TMap> map, IFormatProvider? provider) where TMap : ValuePipeline<int, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.Provider = provider;
        AddNumberMapperIfMissing(map.Pipeline);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the number styles to use when mapping the value of a cell to a <see cref="int"/> or null.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="style">The number styles to use when parsing a cell's text.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithNumberStyle<TMap>(this ValuePipeline<int?, TMap> map, NumberStyles style) where TMap : ValuePipeline<int?, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.NumberStyle = style;
        AddNumberMapperIfMissing(map.Pipeline);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the format provider to use when mapping the value of a cell to a <see cref="int"/> or null.
    /// This is useful for mapping columns where cultures differ.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider to use when parsing a cell's text.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormatProvider<TMap>(this ValuePipeline<int?, TMap> map, IFormatProvider? provider) where TMap : ValuePipeline<int?, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.Provider = provider;
        AddNumberMapperIfMissing(map.Pipeline);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the number styles to use when mapping the value of a cell to a <see cref="uint"/>.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="style">The number styles to use when parsing a cell's text.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithNumberStyle<TMap>(this ValuePipeline<uint, TMap> map, NumberStyles style) where TMap : ValuePipeline<uint, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.NumberStyle = style;
        AddNumberMapperIfMissing(map.Pipeline);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the format provider to use when mapping the value of a cell to a <see cref="uint"/>.
    /// This is useful for mapping columns where cultures differ.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider to use when parsing a cell's text.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormatProvider<TMap>(this ValuePipeline<uint, TMap> map, IFormatProvider? provider) where TMap : ValuePipeline<uint, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.Provider = provider;
        AddNumberMapperIfMissing(map.Pipeline);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the number styles to use when mapping the value of a cell to a <see cref="uint"/> or null.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="style">The number styles to use when parsing a cell's text.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithNumberStyle<TMap>(this ValuePipeline<uint?, TMap> map, NumberStyles style) where TMap : ValuePipeline<uint?, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.NumberStyle = style;
        AddNumberMapperIfMissing(map.Pipeline);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the format provider to use when mapping the value of a cell to a <see cref="uint"/> or null.
    /// This is useful for mapping columns where cultures differ.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider to use when parsing a cell's text.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormatProvider<TMap>(this ValuePipeline<uint?, TMap> map, IFormatProvider? provider) where TMap : ValuePipeline<uint?, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.Provider = provider;
        AddNumberMapperIfMissing(map.Pipeline);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the number styles to use when mapping the value of a cell to a <see cref="long"/>.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="style">The number styles to use when parsing a cell's text.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithNumberStyle<TMap>(this ValuePipeline<long, TMap> map, NumberStyles style) where TMap : ValuePipeline<long, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.NumberStyle = style;
        AddNumberMapperIfMissing(map.Pipeline);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the format provider to use when mapping the value of a cell to a <see cref="long"/>.
    /// This is useful for mapping columns where cultures differ.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider to use when parsing a cell's text.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormatProvider<TMap>(this ValuePipeline<long, TMap> map, IFormatProvider? provider) where TMap : ValuePipeline<long, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.Provider = provider;
        AddNumberMapperIfMissing(map.Pipeline);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the number styles to use when mapping the value of a cell to a <see cref="long"/> or null.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="style">The number styles to use when parsing a cell's text.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithNumberStyle<TMap>(this ValuePipeline<long?, TMap> map, NumberStyles style) where TMap : ValuePipeline<long?, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.NumberStyle = style;
        AddNumberMapperIfMissing(map.Pipeline);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the format provider to use when mapping the value of a cell to a <see cref="long"/> or null.
    /// This is useful for mapping columns where cultures differ.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider to use when parsing a cell's text.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormatProvider<TMap>(this ValuePipeline<long?, TMap> map, IFormatProvider? provider) where TMap : ValuePipeline<long?, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.Provider = provider;
        AddNumberMapperIfMissing(map.Pipeline);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the number styles to use when mapping the value of a cell to a <see cref="ulong"/>.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="style">The number styles to use when parsing a cell's text.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithNumberStyle<TMap>(this ValuePipeline<ulong, TMap> map, NumberStyles style) where TMap : ValuePipeline<ulong, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.NumberStyle = style;
        AddNumberMapperIfMissing(map.Pipeline);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the format provider to use when mapping the value of a cell to a <see cref="ulong"/>.
    /// This is useful for mapping columns where cultures differ.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider to use when parsing a cell's text.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormatProvider<TMap>(this ValuePipeline<ulong, TMap> map, IFormatProvider? provider) where TMap : ValuePipeline<ulong, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.Provider = provider;
        AddNumberMapperIfMissing(map.Pipeline);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the number styles to use when mapping the value of a cell to a <see cref="ulong"/> or null.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="style">The number styles to use when parsing a cell's text.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithNumberStyle<TMap>(this ValuePipeline<ulong?, TMap> map, NumberStyles style) where TMap : ValuePipeline<ulong?, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.NumberStyle = style;
        AddNumberMapperIfMissing(map.Pipeline);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the format provider to use when mapping the value of a cell to a <see cref="ulong"/> or null.
    /// This is useful for mapping columns where cultures differ.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider to use when parsing a cell's text.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormatProvider<TMap>(this ValuePipeline<ulong?, TMap> map, IFormatProvider? provider) where TMap : ValuePipeline<ulong?, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.Provider = provider;
        AddNumberMapperIfMissing(map.Pipeline);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the number styles to use when mapping the value of a cell to a <see cref="float"/>.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="style">The number styles to use when parsing a cell's text.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithNumberStyle<TMap>(this ValuePipeline<float, TMap> map, NumberStyles style) where TMap : ValuePipeline<float, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.NumberStyle = style;
        AddNumberMapperIfMissing(map.Pipeline);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the format provider to use when mapping the value of a cell to a <see cref="float"/>.
    /// This is useful for mapping columns where cultures differ.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider to use when parsing a cell's text.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormatProvider<TMap>(this ValuePipeline<float, TMap> map, IFormatProvider? provider) where TMap : ValuePipeline<float, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.Provider = provider;
        AddNumberMapperIfMissing(map.Pipeline);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the number styles to use when mapping the value of a cell to a <see cref="float"/> or null.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="style">The number styles to use when parsing a cell's text.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithNumberStyle<TMap>(this ValuePipeline<float?, TMap> map, NumberStyles style) where TMap : ValuePipeline<float?, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.NumberStyle = style;
        AddNumberMapperIfMissing(map.Pipeline);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the format provider to use when mapping the value of a cell to a <see cref="float"/> or null.
    /// This is useful for mapping columns where cultures differ.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider to use when parsing a cell's text.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormatProvider<TMap>(this ValuePipeline<float?, TMap> map, IFormatProvider? provider) where TMap : ValuePipeline<float?, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.Provider = provider;
        AddNumberMapperIfMissing(map.Pipeline);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the number styles to use when mapping the value of a cell to a <see cref="double"/>.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="style">The number styles to use when parsing a cell's text.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithNumberStyle<TMap>(this ValuePipeline<double, TMap> map, NumberStyles style) where TMap : ValuePipeline<double, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.NumberStyle = style;
        AddNumberMapperIfMissing(map.Pipeline);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the format provider to use when mapping the value of a cell to a <see cref="double"/>.
    /// This is useful for mapping columns where cultures differ.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider to use when parsing a cell's text.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormatProvider<TMap>(this ValuePipeline<double, TMap> map, IFormatProvider? provider) where TMap : ValuePipeline<double, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.Provider = provider;
        AddNumberMapperIfMissing(map.Pipeline);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the number styles to use when mapping the value of a cell to a <see cref="double"/> or null.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="style">The number styles to use when parsing a cell's text.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithNumberStyle<TMap>(this ValuePipeline<double?, TMap> map, NumberStyles style) where TMap : ValuePipeline<double?, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.NumberStyle = style;
        AddNumberMapperIfMissing(map.Pipeline);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the format provider to use when mapping the value of a cell to a <see cref="double"/> or null.
    /// This is useful for mapping columns where cultures differ.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider to use when parsing a cell's text.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormatProvider<TMap>(this ValuePipeline<double?, TMap> map, IFormatProvider? provider) where TMap : ValuePipeline<double?, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.Provider = provider;
        AddNumberMapperIfMissing(map.Pipeline);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the number styles to use when mapping the value of a cell to a <see cref="decimal"/>.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="style">The number styles to use when parsing a cell's text.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithNumberStyle<TMap>(this ValuePipeline<decimal, TMap> map, NumberStyles style) where TMap : ValuePipeline<decimal, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.NumberStyle = style;
        AddNumberMapperIfMissing(map.Pipeline);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the format provider to use when mapping the value of a cell to a <see cref="decimal"/>.
    /// This is useful for mapping columns where cultures differ.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider to use when parsing a cell's text.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormatProvider<TMap>(this ValuePipeline<decimal, TMap> map, IFormatProvider? provider) where TMap : ValuePipeline<decimal, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.Provider = provider;
        AddNumberMapperIfMissing(map.Pipeline);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the number styles to use when mapping the value of a cell to a <see cref="decimal"/> or null.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="style">The number styles to use when parsing a cell's text.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithNumberStyle<TMap>(this ValuePipeline<decimal?, TMap> map, NumberStyles style) where TMap : ValuePipeline<decimal?, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.NumberStyle = style;
        AddNumberMapperIfMissing(map.Pipeline);
        return (TMap)map;
    }

    /// <summary>
    /// Specifies the format provider to use when mapping the value of a cell to a <see cref="decimal"/> or null.
    /// This is useful for mapping columns where cultures differ.
    /// </summary>
    /// <typeparam name="TMap">The type of the map.</typeparam>
    /// <param name="map">The map to use.</param>
    /// <param name="provider">The format provider to use when parsing a cell's text.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public static TMap WithFormatProvider<TMap>(this ValuePipeline<decimal?, TMap> map, IFormatProvider? provider) where TMap : ValuePipeline<decimal?, TMap>
    {
        ThrowHelpers.ThrowIfNull(map, nameof(map));
        map.Pipeline.Options.Provider = provider;
        AddNumberMapperIfMissing(map.Pipeline);
        return (TMap)map;
    }
}

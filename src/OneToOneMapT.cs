using System.Collections.Concurrent;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Fallbacks;

namespace ExcelMapper;

/// <summary>
/// Maps a member to an Excel column of type T with a one-to-one relationship.
/// </summary>
/// <typeparam name="T">The type of the member.</typeparam>
public class OneToOneMap<T> : IOneToOneMap, IValuePipeline<T>, IBindableMap
{
    /// <summary>
    /// Constructs a one-to-one map with the specified cell reader factory.
    /// </summary>
    /// <param name="readerFactory">The factory used to create cell readers for the mapped column.</param>
    public OneToOneMap(ICellReaderFactory readerFactory)
    {
        ThrowHelpers.ThrowIfNull(readerFactory, nameof(readerFactory));
        _readerFactory = readerFactory;
    }

    private ICellReaderFactory _readerFactory;

    /// <inheritdoc />
    public ICellReaderFactory ReaderFactory
    {
        get => _readerFactory;
        set
        {
            ThrowHelpers.ThrowIfNull(value, nameof(value));
            _readerFactory = value;
            _factoryCache.Clear();
            ReaderVersion++;
        }
    }

    /// <inheritdoc />
    public bool Optional { get; set; }

    /// <inheritdoc />
    public bool PreserveFormatting { get; set; }

    /// <inheritdoc />
    public IValuePipeline Pipeline { get; } = new ValuePipeline<T>();

    /// <inheritdoc />
    public IList<ICellTransformer> Transformers => Pipeline.Transformers;

    /// <inheritdoc />
    public IList<ICellMapper> Mappers => Pipeline.Mappers;

    /// <inheritdoc />
    public IFallbackItem? EmptyFallback
    {
        get => Pipeline.EmptyFallback;
        set => Pipeline.EmptyFallback = value;
    }

    /// <inheritdoc />
    public IFallbackItem? InvalidFallback
    {
        get => Pipeline.InvalidFallback;
        set => Pipeline.InvalidFallback = value;
    }

    private readonly PerSheetCache<ICellReader> _factoryCache = new();

    /// <inheritdoc />
    public bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, [NotNullWhen(true)] out object? result)
    {
        if (TryGetValue(GetCellReader(sheet), sheet, rowIndex, reader, member, out var value))
        {
            result = value!;
            return true;
        }

        result = null;
        return false;
    }

    /// <summary>
    /// Incremented when <see cref="ReaderFactory"/> changes, so that readers resolved for a sheet
    /// (see <see cref="BoundOneToOneMember{T}"/>) are resolved again.
    /// </summary>
    internal int ReaderVersion { get; private set; }

    internal ICellReader? GetCellReader(ExcelSheet sheet)
        => _factoryCache.GetOrAdd(sheet, _readerFactory, static (s, factory) => factory.GetCellReader(s));

    /// <summary>
    /// Reads and maps the cell, returning the value as <typeparamref name="T"/> so that value types
    /// aren't boxed again on the way to the member.
    /// </summary>
    internal bool TryGetValue(ICellReader? cellReader, ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, out T value)
    {
        if (cellReader == null || !cellReader.TryGetValue(reader, PreserveFormatting, out var readResult))
        {
            if (Optional)
            {
                value = default!;
                return false;
            }

            throw ExcelMappingException.CreateForNoSuchColumn(sheet, rowIndex, _readerFactory, member);
        }

        if (GetConverters() is { } converters)
        {
            value = Convert(converters, sheet, rowIndex, readResult, member);
            return true;
        }

        var result = ValuePipeline.GetPropertyValue(Pipeline, sheet, rowIndex, readResult, PreserveFormatting, member);
        value = ValuePipeline.CastValue<T>(result, sheet, rowIndex, readResult.ColumnIndex, member)!;
        return true;
    }

    private sealed class Converters(int version, ICellConverter<T>[]? items)
    {
        public int Version { get; } = version;

        public ICellConverter<T>[]? Items { get; } = items;
    }

    private Converters? _converters;

    /// <summary>
    /// Gets the mappers as typed converters, or null if any of them isn't one (for example a custom
    /// mapper), in which case values go through the pipeline as objects. Rebuilt when the mappers change.
    /// </summary>
    private ICellConverter<T>[]? GetConverters()
    {
        var mappers = (NonNullCollection<ICellMapper>)Pipeline.Mappers;
        var converters = _converters;
        if (converters is null || converters.Version != mappers.Version)
        {
            converters = new Converters(mappers.Version, CreateConverters(mappers));
            _converters = converters;
        }

        return converters.Items;
    }

    private static ICellConverter<T>[]? CreateConverters(NonNullCollection<ICellMapper> mappers)
    {
        var converters = new ICellConverter<T>[mappers.Count];
        for (int i = 0; i < converters.Length; i++)
        {
            if (mappers[i] is not ICellConverter<T> converter)
            {
                return null;
            }

            converters[i] = converter;
        }

        return converters;
    }

    /// <summary>
    /// Maps a cell through the typed converters, with the same rules as
    /// <see cref="ValuePipeline.GetPropertyValue"/>: transformers, the empty fallback, then the
    /// converters (a success stops, a success that can be replaced and an invalid value are kept
    /// until a later result replaces them), then the invalid fallback.
    /// </summary>
    private T Convert(ICellConverter<T>[] converters, ExcelSheet sheet, int rowIndex, ReadCellResult readResult, MemberInfo? member)
    {
        var pipeline = Pipeline;
        readResult = ValuePipeline.ApplyTransformers(pipeline, sheet, rowIndex, readResult, PreserveFormatting);
        if (readResult.IsEmpty() && pipeline.EmptyFallback != null)
        {
            return ValuePipeline.CastValue<T>(pipeline.EmptyFallback.PerformFallback(sheet, rowIndex, readResult, null, member), sheet, rowIndex, readResult.ColumnIndex, member)!;
        }

        var state = CellConversion.NotHandled;
        T result = default!;
        var invalidIndex = -1;
        for (int i = 0; i < converters.Length; i++)
        {
            switch (converters[i].Convert(readResult, out var value))
            {
                case CellConversion.Success:
                    return value;
                case CellConversion.SuccessUnlessReplaced:
                    state = CellConversion.SuccessUnlessReplaced;
                    result = value;
                    break;
                case CellConversion.Invalid:
                    state = CellConversion.Invalid;
                    invalidIndex = i;
                    break;
            }
        }

        if (state == CellConversion.SuccessUnlessReplaced)
        {
            return result;
        }

        if (pipeline.InvalidFallback is { } invalidFallback)
        {
            // Create the exception with the mapper's own Map, only if the fallback uses it.
            var exception = state == CellConversion.Invalid && ValuePipeline.FallbackUsesException(invalidFallback)
                ? converters[invalidIndex].Map(readResult).Exception
                : null;
            return ValuePipeline.CastValue<T>(invalidFallback.PerformFallback(sheet, rowIndex, readResult, exception, member), sheet, rowIndex, readResult.ColumnIndex, member)!;
        }

        return ValuePipeline.CastValue<T>(null, sheet, rowIndex, readResult.ColumnIndex, member)!;
    }

    /// <summary>
    /// Reads and maps the cell as an object, for a constructor argument. A value that is already a
    /// <typeparamref name="T"/> is returned as is, rather than unboxed and boxed again.
    /// </summary>
    internal bool TryGetObjectValue(ICellReader? cellReader, ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, out object? value)
    {
        if (cellReader == null || !cellReader.TryGetValue(reader, PreserveFormatting, out var readResult))
        {
            if (Optional)
            {
                value = null;
                return false;
            }

            throw ExcelMappingException.CreateForNoSuchColumn(sheet, rowIndex, _readerFactory, member);
        }

        // Constructor arguments are objects, so map through the pipeline as objects: mappers can
        // return cached boxes (such as for bool) that the typed path would box again.
        var result = ValuePipeline.GetPropertyValue(Pipeline, sheet, rowIndex, readResult, PreserveFormatting, member);
        value = result is T && !ReferenceEquals(result, ThrowFallback.ErrorRecorded)
            ? result
            : ValuePipeline.CastValue<T>(result, sheet, rowIndex, readResult.ColumnIndex, member);
        return true;
    }

    /// <inheritdoc/>
    BoundMember IBindableMap.Bind(ExcelPropertyMap property, ExcelSheet sheet) => new BoundOneToOneMember<T>(this, property, sheet);

    /// <summary>
    /// Specifies a fixed fallback to be used if the value of a cell is empty or cannot be mapped.
    /// </summary>
    /// <typeparam name="T">The type of the map.</typeparam>
    /// <param name="value">The value that will be assigned to the property or field if the value of a cell is empty or cannot be mapped.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public OneToOneMap<T> WithValueFallback(T? value)
        => this.WithFallbackItem(new FixedValueFallback(value));

    /// <summary>
    /// Specifies a fixed fallback to be used if the value of a cell is empty.
    /// </summary>
    /// <typeparam name="T">The type of the map.</typeparam>
    /// <param name="fallbackValue">The value that will be assigned to the property or field if the value of a cell is empty.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public OneToOneMap<T> WithEmptyFallback(T? fallbackValue)
        => this.WithEmptyFallbackItem(new FixedValueFallback(fallbackValue));

    /// <summary>
    /// Specifies a fixed fallback to be used if the value of a cell cannot be mapped.
    /// </summary>
    /// <typeparam name="T">The type of the map.</typeparam>
    /// <param name="fallbackValue">The value that will be assigned to the property or field if the value of a cell cannot be mapped.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public OneToOneMap<T> WithInvalidFallback(T? fallbackValue)
        => this.WithInvalidFallbackItem(new FixedValueFallback(fallbackValue));
}

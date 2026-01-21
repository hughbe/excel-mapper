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
public class OneToOneMap<T> : IOneToOneMap, IValuePipeline<T>
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

    private readonly ConcurrentDictionary<ExcelSheet, ICellReader?> _factoryCache = new();

    /// <inheritdoc />
    public bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, [NotNullWhen(true)] out object? result)
    {
        var cellReader = _factoryCache.GetOrAdd(sheet, s => _readerFactory.GetCellReader(s));

        if (cellReader == null || !cellReader.TryGetValue(reader, PreserveFormatting, out var readResult))
        {
            if (Optional)
            {
                result = default;
                return false;
            }

            throw ExcelMappingException.CreateForNoSuchColumn(sheet, rowIndex, _readerFactory, member);
        }

        result = (T)ValuePipeline.GetPropertyValue(Pipeline, sheet, rowIndex, readResult, PreserveFormatting, member)!;
        return true;
    }

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

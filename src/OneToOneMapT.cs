using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Abstractions;
using ExcelMapper.Fallbacks;

namespace ExcelMapper;

public class OneToOneMap<T> : IOneToOneMap, IValuePipeline<T>
{
    public OneToOneMap(ICellReaderFactory readerFactory)
    {
        _readerFactory = readerFactory ?? throw new ArgumentNullException(nameof(readerFactory));
    }

    private ICellReaderFactory _readerFactory;

    public ICellReaderFactory ReaderFactory
    {
        get => _readerFactory;
        set => _readerFactory = value ?? throw new ArgumentNullException(nameof(value));
    }

    /// <inheritdoc />
    public bool Optional { get; set; }

    /// <inheritdoc />
    public bool PreserveFormatting { get; set; }

    public ValuePipeline<T> Pipeline { get; } = new ValuePipeline<T>();

    private readonly ConcurrentDictionary<ExcelSheet, ICellReader?> _factoryCache = new();

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

    public IReadOnlyList<ICellTransformer> CellValueTransformers => Pipeline.CellValueTransformers;

    public IReadOnlyList<ICellMapper> CellValueMappers => Pipeline.CellValueMappers;

    public IFallbackItem? EmptyFallback
    {
        get => Pipeline.EmptyFallback;
        set => Pipeline.EmptyFallback = value;
    }

    public IFallbackItem? InvalidFallback
    {
        get => Pipeline.InvalidFallback;
        set => Pipeline.InvalidFallback = value;
    }

    public void AddCellValueMapper(ICellMapper mapper) => Pipeline.AddCellValueMapper(mapper);

    public void AddCellValueTransformer(ICellTransformer transformer) => Pipeline.AddCellValueTransformer(transformer);

    public void RemoveCellValueMapper(int index) => Pipeline.RemoveCellValueMapper(index);

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

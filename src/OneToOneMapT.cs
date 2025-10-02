using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Abstractions;
using ExcelMapper.Fallbacks;

namespace ExcelMapper;

public class OneToOneMap<T> : IValuePipeline<T>, IMap
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

    public bool Optional { get; set; }

    public ValuePipeline<T> Pipeline { get; } = new ValuePipeline<T>();

    private readonly Dictionary<ExcelSheet, ICellReader?> _factoryCache = [];

    public bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, [NotNullWhen(true)] out object? result)
    {
        if (!_factoryCache.TryGetValue(sheet, out ICellReader? cellReader))
        {
            cellReader = _readerFactory.GetReader(sheet);
            _factoryCache.Add(sheet, cellReader);
        }

        if (cellReader == null || !cellReader.TryGetValue(reader, out ReadCellResult readResult))
        {
            if (Optional)
            {
                result = default;
                return false;
            }

            throw new ExcelMappingException($"Could not read value for {member?.Name}", sheet, rowIndex, -1);
        }

        result = (T)ValuePipeline.GetPropertyValue(Pipeline, sheet, rowIndex, readResult, member)!;
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
    /// <typeparam name="T">The type of the property map.</typeparam>
    /// <param name="defaultValue">The value that will be assigned to the property or field if the value of a cell is empty or cannot be mapped.</param>
    /// <returns>The property map on which this method was invoked.</returns>
    public OneToOneMap<T> WithValueFallback(T? defaultValue)
    {
        return this
            .WithEmptyFallback(defaultValue)
            .WithInvalidFallback(defaultValue);
    }

    /// <summary>
    /// Specifies a fixed fallback to be used if the value of a cell is empty.
    /// </summary>
    /// <typeparam name="T">The type of the property map.</typeparam>
    /// <param name="fallbackValue">The value that will be assigned to the property or field if the value of a cell is empty.</param>
    /// <returns>The property map on which this method was invoked.</returns>
    public OneToOneMap<T> WithEmptyFallback(T? fallbackValue)
    {
        return this.WithEmptyFallbackItem(new FixedValueFallback(fallbackValue));
    }

    /// <summary>
    /// Specifies a fixed fallback to be used if the value of a cell cannot be mapped.
    /// </summary>
    /// <typeparam name="T">The type of the property map.</typeparam>
    /// <param name="fallbackValue">The value that will be assigned to the property or field if the value of a cell cannot be mapped.</param>
    /// <returns>The property map on which this method was invoked.</returns>
    public OneToOneMap<T> WithInvalidFallback(T? fallbackValue)
    {
        return this.WithInvalidFallbackItem(new FixedValueFallback(fallbackValue));
    }
}

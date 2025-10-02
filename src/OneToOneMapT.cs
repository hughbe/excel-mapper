using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Abstractions;

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

        result = (T?)ValuePipeline.GetPropertyValue(Pipeline, sheet, rowIndex, readResult, member);
        return result != null;
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
}

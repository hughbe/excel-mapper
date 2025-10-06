using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Abstractions;
using ExcelMapper.Readers;

namespace ExcelMapper;

public delegate object CreateDictionaryFactory<TElement>(IEnumerable<KeyValuePair<string, TElement>> elements);

/// <summary>
/// A map that reads one or more values from one or more cells and maps these values to the type of the
/// property or field. This is used to map IDictionary properties and fields.
/// </summary>
/// <typeparam name="TElement">The element type of the IDictionary property or field.</typeparam>
public class ManyToOneDictionaryMap<TElement> : IMap
{
    /// <summary>
    /// Constructs a map reads one or more values from one or more cells and maps these values as element
    /// contained by the property or field.
    /// </summary>
    /// <param name="valuePipeline">The map that maps the value of a single cell to an object of the element type of the property or field.</param>
    public ManyToOneDictionaryMap(ICellsReaderFactory readerFactory, IValuePipeline<TElement> valuePipeline, CreateDictionaryFactory<TElement> createDictionaryFactory)
    {
        _readerFactory = readerFactory ?? throw new ArgumentNullException(nameof(readerFactory));
        ValuePipeline = valuePipeline ?? throw new ArgumentNullException(nameof(valuePipeline));
        CreateDictionaryFactory = createDictionaryFactory ?? throw new ArgumentNullException(nameof(createDictionaryFactory));
    }

    /// <summary>
    /// Gets the map that maps the value of a single cell to an object of the element type of the property
    /// or field.
    /// </summary>
    public IValuePipeline<TElement> ValuePipeline { get; private set; }
    
    public bool Optional { get; set; }

    public bool PreserveFormatting { get; set; }

    /// <summary>
    /// Gets the reader that reads one or more values from one or more cells used to map each
    /// element of the property or field.
    /// </summary>
    private ICellsReaderFactory _readerFactory;

    public ICellsReaderFactory ReaderFactory
    {
        get => _readerFactory;
        set => _readerFactory = value ?? throw new ArgumentNullException(nameof(value));
    }

    public CreateDictionaryFactory<TElement> CreateDictionaryFactory { get; }
    
    private readonly Dictionary<ExcelSheet, ICellsReader?> _factoryCache = [];

    public bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, [NotNullWhen(true)] out object? result)
    {
        if (sheet == null)
        {
            throw new ArgumentNullException(nameof(sheet));
        }
        if (sheet.Heading == null)
        {
            throw new ExcelMappingException("The sheet \"{sheet.Name}\" does not have a heading. Use a column index map instead.");
        }
        if (!_factoryCache.TryGetValue(sheet, out var cellsReader))
        {
            cellsReader = _readerFactory.GetCellsReader(sheet);
            _factoryCache.Add(sheet, cellsReader);
        }
        
        if (cellsReader == null || !cellsReader.TryGetValues(reader, PreserveFormatting, out IEnumerable<ReadCellResult>? valueResults))
        {
            if (Optional)
            {
                result = default;
                return false;
            }

            throw new ExcelMappingException($"Could not read value for \"{member?.Name}\"", sheet, rowIndex, -1);
        }

        var valueResultsList = valueResults.ToList();

        var values = new List<TElement>();
        foreach (ReadCellResult valueResult in valueResultsList)
        {
            // Discarding nullability check because it may be indended to be this way (T may be nullable)
            TElement value = (TElement)ExcelMapper.ValuePipeline.GetPropertyValue(ValuePipeline, sheet, rowIndex, valueResult, PreserveFormatting, member)!;
            values.Add(value);
        }

        var heading = sheet.Heading;
        var keys = valueResultsList.Select(r => heading.GetColumnName(r.ColumnIndex));
        var elements = keys.Zip(values, (key, keyValue) => new KeyValuePair<string, TElement>(key, keyValue));
        result = CreateDictionaryFactory(elements);
        return true;
    }

    /// <summary>
    /// Sets the reader of the property map to read the values of one or more cells contained
    /// in the columns with the given names.
    /// </summary>
    /// <param name="columnNames">The name of each column to read.</param>
    /// <returns>The property map that invoked this method.</returns>
    public ManyToOneDictionaryMap<TElement> WithColumnNames(params string[] columnNames)
    {
        ReaderFactory = new ColumnNamesReaderFactory(columnNames);
        return this;
    }

    /// <summary>
    /// Sets the reader of the property map to read the values of one or more cells contained
    /// in the columns with the given names.
    /// </summary>
    /// <param name="columnNames">The name of each column to read.</param>
    /// <returns>The property map that invoked this method.</returns>
    public ManyToOneDictionaryMap<TElement> WithColumnNames(IEnumerable<string> columnNames)
    {
        if (columnNames == null)
        {
            throw new ArgumentNullException(nameof(columnNames));
        }

        return WithColumnNames([.. columnNames]);
    }

    /// <summary>
    /// Sets the reader of the property map to read the values of one or more cells contained
    /// in the columns matching the result of IExcelColumnMatcher.ColumnMatches.
    /// </summary>
    /// <param name="matcher">The matcher of each column to read.</param>
    /// <returns>The property map that invoked this method.</returns>
    public ManyToOneDictionaryMap<TElement> WithColumnsMatching(IExcelColumnMatcher matcher)
    {
        ReaderFactory = new ColumnsMatchingReaderFactory(matcher);
        return this;
    }

    /// <summary>
    /// Sets the reader of the property map to read the values of one or more cells contained
    /// in the columns with the given zero-based indices.
    /// </summary>
    /// <param name="columnIndices">The zero-based index of each column to read.</param>
    /// <returns>The property map that invoked this method.</returns>
    public ManyToOneDictionaryMap<TElement> WithColumnIndices(params int[] columnIndices)
    {
        ReaderFactory = new ColumnIndicesReaderFactory(columnIndices);
        return this;
    }

    /// <summary>
    /// Sets the reader of the property map to read the values of one or more cells contained
    /// in the columns with the given zero-based indices.
    /// </summary>
    /// <param name="columnIndices">The zero-based index of each column to read.</param>
    /// <returns>The property map that invoked this method.</returns>
    public ManyToOneDictionaryMap<TElement> WithColumnIndices(IEnumerable<int> columnIndices)
    {
        if (columnIndices == null)
        {
            throw new ArgumentNullException(nameof(columnIndices));
        }
        
        return WithColumnIndices([.. columnIndices]);
    }

    /// <summary>
    /// Makes the reader of the property map peserve formatting when reading string values.
    /// </summary>
    /// <returns>The property map on which this method was invoked.</returns>
    public ManyToOneDictionaryMap<TElement> MakeOptional()
    {
        Optional = true;
        return this;
    }

    /// <summary>
    /// Makes the reader of the property map optional. For example, if the column doesn't exist
    /// or the index is invalid, an exception will not be thrown.
    /// </summary>
    /// <returns>The property map on which this method was invoked.</returns>
    public ManyToOneDictionaryMap<TElement> MakePreserveFormatting()
    {
        PreserveFormatting = true;
        return this;
    }

    /// <summary>
    /// Sets the map that maps the value of a single cell to an object of the element type of the property
    /// or field.
    /// </summary>
    /// <param name="valueMap">The pipeline that maps the value of a single cell to an object of the element type of the property
    /// or field.</param>
    /// <returns>The property map that invoked this method.</returns>
    public ManyToOneDictionaryMap<TElement> WithValueMap(Func<IValuePipeline<TElement>, IValuePipeline<TElement>> valueMap)
    {
        if (valueMap == null)
        {
            throw new ArgumentNullException(nameof(valueMap));
        }

        ValuePipeline = valueMap(ValuePipeline) ?? throw new ArgumentNullException(nameof(valueMap));
        return this;
    }
}

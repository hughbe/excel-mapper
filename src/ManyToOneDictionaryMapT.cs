using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Abstractions;

namespace ExcelMapper;

/// <summary>
/// A map that reads one or more values from one or more cells and maps these values to the type of the
/// property or field. This is used to map IDictionary properties and fields.
/// </summary>
/// <typeparam name="TValue">The value type of the IDictionary property or field.</typeparam>
public class ManyToOneDictionaryMap<TValue> : IManyToOneMap
{
    /// <summary>
    /// Constructs a map reads one or more values from one or more cells and maps these values as element
    /// contained by the property or field.
    /// </summary>
    /// <param name="valuePipeline">The map that maps the value of a single cell to an object of the element type of the property or field.</param>
    public ManyToOneDictionaryMap(ICellsReaderFactory readerFactory, IValuePipeline<TValue> valuePipeline, IDictionaryFactory<TValue> dictionaryFactory)
    {
        _readerFactory = readerFactory ?? throw new ArgumentNullException(nameof(readerFactory));
        ValuePipeline = valuePipeline ?? throw new ArgumentNullException(nameof(valuePipeline));
        DictionaryFactory = dictionaryFactory ?? throw new ArgumentNullException(nameof(dictionaryFactory));
    }

    /// <summary>
    /// Gets the map that maps the value of a single cell to an object of the element type of the property
    /// or field.
    /// </summary>
    public IValuePipeline<TValue> ValuePipeline { get; private set; }
    
    /// <inheritdoc />
    public bool Optional { get; set; }

    /// <inheritdoc />
    public bool PreserveFormatting { get; set; }

    private ICellsReaderFactory _readerFactory;

    /// <inheritdoc />
    public ICellsReaderFactory ReaderFactory
    {
        get => _readerFactory;
        set => _readerFactory = value ?? throw new ArgumentNullException(nameof(value));
    }

    /// <summary>
    /// The factory for creating and adding elements to the list.
    /// </summary>
    public IDictionaryFactory<TValue> DictionaryFactory { get; }
    
    private readonly Dictionary<ExcelSheet, ICellsReader?> _factoryCache = [];

    public bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, [NotNullWhen(true)] out object? value)
    {
        if (sheet == null)
        {
            throw new ArgumentNullException(nameof(sheet));
        }
        if (sheet.Heading == null)
        {
            throw new ExcelMappingException($"The sheet \"{sheet.Name}\" does not have a heading. Use a column index map instead.");
        }
        if (!_factoryCache.TryGetValue(sheet, out var cellsReader))
        {
            cellsReader = _readerFactory.GetCellsReader(sheet);
            _factoryCache.Add(sheet, cellsReader);
        }

        if (cellsReader == null || !cellsReader.TryGetValues(reader, PreserveFormatting, out var valueResults))
        {
            if (Optional)
            {
                value = default;
                return false;
            }

            throw new ExcelMappingException($"Could not read value for \"{member?.Name}\"", sheet, rowIndex, -1);
        }

        DictionaryFactory.Begin(valueResults.Count());
        try
        {
            foreach (var valueResult in valueResults)
            {
                var elementKey = sheet.Heading.GetColumnName(valueResult.ColumnIndex);
                var elementValue = (TValue)ExcelMapper.ValuePipeline.GetPropertyValue(ValuePipeline, sheet, rowIndex, valueResult, PreserveFormatting, member)!;
                DictionaryFactory.Add(elementKey, elementValue);
            }

            value = DictionaryFactory.End();
            return true;
        }
        finally
        {
            DictionaryFactory.Reset();
        }
    }

    /// <summary>
    /// Sets the map that maps the value of a single cell to an object of the element type of the property
    /// or field.
    /// </summary>
    /// <param name="valueMap">The pipeline that maps the value of a single cell to an object of the element type of the property
    /// or field.</param>
    /// <returns>The map that invoked this method.</returns>
    public ManyToOneDictionaryMap<TValue> WithValueMap(Func<IValuePipeline<TValue>, IValuePipeline<TValue>> valueMap)
    {
        if (valueMap == null)
        {
            throw new ArgumentNullException(nameof(valueMap));
        }

        ValuePipeline = valueMap(ValuePipeline) ?? throw new ArgumentNullException(nameof(valueMap));
        return this;
    }
}

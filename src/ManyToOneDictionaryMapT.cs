using System.Collections.Concurrent;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Reflection;
using ExcelDataReader;

namespace ExcelMapper;

/// <summary>
/// A map that reads one or more values from one or more cells and maps these values to the type of the
/// property or field. This is used to map IDictionary properties and fields.
/// </summary>
/// <typeparam name="TKey">The key type of the IDictionary property or field.</typeparam>
/// <typeparam name="TValue">The value type of the IDictionary property or field.</typeparam>
public class ManyToOneDictionaryMap<TKey, TValue> : IManyToOneMap where TKey : notnull
{
    /// <summary>
    /// Constructs a map reads one or more values from one or more cells and maps these values as element
    /// contained by the property or field.
    /// </summary>
    /// <param name="valuePipeline">The map that maps the value of a single cell to an object of the element type of the property or field.</param>
    public ManyToOneDictionaryMap(ICellsReaderFactory readerFactory, IDictionaryFactory<TKey, TValue> dictionaryFactory)
    {
        ArgumentNullException.ThrowIfNull(readerFactory);
        ArgumentNullException.ThrowIfNull(dictionaryFactory);

        _readerFactory = readerFactory;
        DictionaryFactory = dictionaryFactory;
    }

    /// <summary>
    /// Gets the map that maps the value of a single cell to an object of the element type of the property
    /// or field.
    /// </summary>
    public IValuePipeline Pipeline { get; private set; } = new ValuePipeline<TValue>();
    
    /// <inheritdoc />
    public bool Optional { get; set; }

    /// <inheritdoc />
    public bool PreserveFormatting { get; set; }

    private ICellsReaderFactory _readerFactory;

    /// <inheritdoc />
    public ICellsReaderFactory ReaderFactory
    {
        get => _readerFactory;
        set
        {
            ArgumentNullException.ThrowIfNull(value);
            _readerFactory = value;
        }
    }

    /// <summary>
    /// The factory for creating and adding elements to the list.
    /// </summary>
    public IDictionaryFactory<TKey, TValue> DictionaryFactory { get; }
    
    private readonly ConcurrentDictionary<ExcelSheet, ICellsReader?> _factoryCache = new();

    public bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, [NotNullWhen(true)] out object? value)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        if (sheet.Heading == null)
        {
            throw new ExcelMappingException($"The sheet \"{sheet.Name}\" does not have a heading. Use a column index map instead.");
        }
        
        var cellsReader = _factoryCache.GetOrAdd(sheet, s => _readerFactory.GetCellsReader(s));

        if (cellsReader == null || !cellsReader.TryGetValues(reader, PreserveFormatting, out var valueResults))
        {
            if (Optional)
            {
                value = default;
                return false;
            }

            throw ExcelMappingException.CreateForNoSuchColumn(sheet, rowIndex, _readerFactory, member);
        }

        DictionaryFactory.Begin(valueResults.Count());
        try
        {
            foreach (var valueResult in valueResults)
            {
                var elementKey = sheet.Heading.GetColumnName(valueResult.ColumnIndex);
                // Convert the string key to TKey
                TKey? convertedKey = (TKey?)Convert.ChangeType(elementKey, typeof(TKey));
                if (convertedKey != null)
                {
                    var elementValue = (TValue)ValuePipeline.GetPropertyValue(Pipeline, sheet, rowIndex, valueResult, PreserveFormatting, member)!;
                    DictionaryFactory.Add(convertedKey, elementValue);
                }
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
    public ManyToOneDictionaryMap<TKey, TValue> WithValueMap(Func<IValuePipeline<TValue>, IValuePipeline<TValue>> valueMap)
    {
        ArgumentNullException.ThrowIfNull(valueMap);

        var result = valueMap((IValuePipeline<TValue>)Pipeline);
        ArgumentNullException.ThrowIfNull(result, nameof(valueMap));
        Pipeline = result;

        return this;
    }
}

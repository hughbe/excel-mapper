using ExcelMapper.Readers;
using System.Text.RegularExpressions;
using System.Collections.Concurrent;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelDataReader;

namespace ExcelMapper;

/// <summary>
/// A map that reads one or more values from one or more cells and maps these values to the type of the
/// property or field. This is used to map IDictionary properties and fields.
/// </summary>
/// <typeparam name="TKey">The key type of the IDictionary property or field.</typeparam>
/// <typeparam name="TValue">The value type of the IDictionary property or field.</typeparam>
public class ManyToOneDictionaryMap<TKey, TValue> : IManyToOneMap, ICellsMap where TKey : notnull
{
    /// <summary>
    /// Constructs a map that reads values from one or more cells and maps them to the values of the
    /// dictionary assigned to the property or field.
    /// </summary>
    /// <param name="readerFactory">Finds the columns to read.</param>
    /// <param name="dictionaryFactory">Creates the dictionary.</param>
    internal ManyToOneDictionaryMap(ICellsReaderFactory readerFactory, IDictionaryFactory<TKey, TValue> dictionaryFactory)
    {
        ThrowHelpers.ThrowIfNull(readerFactory, nameof(readerFactory));
        ThrowHelpers.ThrowIfNull(dictionaryFactory, nameof(dictionaryFactory));

        _readerFactory = readerFactory;
        DictionaryFactory = dictionaryFactory;
    }

    /// <summary>
    /// Gets the map that maps the value of a single cell to an object of the element type of the property
    /// or field.
    /// </summary>
    public ValuePipeline<TValue> Pipeline { get; private set; } = new ValuePipeline<TValue>();

    ValuePipeline IToOneMap.Pipeline => Pipeline;
    
    /// <inheritdoc />
    public bool Optional { get; set; }

    /// <inheritdoc />
    public bool PreserveFormatting { get; set; }

    private ICellsReaderFactory _readerFactory;

    /// <inheritdoc />
    internal ICellsReaderFactory ReaderFactory
    {
        get => _readerFactory;
        set
        {
            ThrowHelpers.ThrowIfNull(value, nameof(value));
            _readerFactory = value;
            ReaderVersion++;
        }
    }

    ICellsReaderFactory IManyToOneMap.ReaderFactory
    {
        get => ReaderFactory;
        set => ReaderFactory = value;
    }

    /// <summary>
    /// The factory for creating and adding elements to the list.
    /// </summary>
    internal IDictionaryFactory<TKey, TValue> DictionaryFactory { get; }

    // Factories build one value at a time, so a map shared between importers (for example
    // a class map instance registered in several importers used concurrently) must not
    // build two values at once.
    private readonly object _buildLock = new();


    /// <inheritdoc/>
    bool IMap.TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, [NotNullWhen(true)] out object? value)
    {
        ThrowHelpers.ThrowIfNull(sheet, nameof(sheet));
        if (sheet.Heading == null)
        {
            throw new ExcelMappingException($"The sheet \"{sheet.Name}\" does not have a heading. Use a column index map instead.");
        }

        return TryGetValue(GetCellsReader(sheet), sheet, rowIndex, reader, member, out value);
    }

    /// <summary>
    /// Incremented when <see cref="ReaderFactory"/> changes, so that readers resolved for a sheet
    /// (see <see cref="BoundCellsMember"/>) are resolved again.
    /// </summary>
    int ICellsMap.ReaderVersion => ReaderVersion;

    internal int ReaderVersion { get; private set; }

    ICellsReader? ICellsMap.GetCellsReader(ExcelSheet sheet) => GetCellsReader(sheet);

    private ICellsReader? GetCellsReader(ExcelSheet sheet)
        => sheet.GetCellsReader(_readerFactory);

    object? ICellsMap.CreateBindingFactory() => (DictionaryFactory as IBindingFactory)?.CreateForBinding();

    bool ICellsMap.TryGetValue(ICellsReader? cellsReader, object? bindingFactory, ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, out object? value)
    {
        if (bindingFactory is IDictionaryFactory<TKey, TValue> factory)
        {
            // The binding's own factory: only used by one sheet, so no lock is needed.
            return TryGetValue(cellsReader, factory, sheet, rowIndex, reader, member, out value);
        }

        return TryGetValue(cellsReader, sheet, rowIndex, reader, member, out value);
    }

    private bool TryGetValue(ICellsReader? cellsReader, ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, [NotNullWhen(true)] out object? value)
    {
        lock (_buildLock)
        {
            return TryGetValue(cellsReader, DictionaryFactory, sheet, rowIndex, reader, member, out value);
        }
    }

    private bool TryGetValue(ICellsReader? cellsReader, IDictionaryFactory<TKey, TValue> factory, ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, [NotNullWhen(true)] out object? value)
    {
        if (sheet.Heading == null)
        {
            throw new ExcelMappingException($"The sheet \"{sheet.Name}\" does not have a heading. Use a column index map instead.");
        }

        
        if (cellsReader == null || !cellsReader.Start(reader, PreserveFormatting, out var count))
        {
            if (Optional)
            {
                value = default;
                return false;
            }

            throw ExcelMappingException.CreateForNoSuchColumn(sheet, rowIndex, _readerFactory, member);
        }

        factory.Begin(count);
        try
        {
            try
            {
                while (cellsReader.TryGetNext(out var valueResult))
                {
                    // Get the column name for the current value
                    var elementKey = sheet.Heading.GetColumnName(valueResult.ColumnIndex);
                    // Convert the string key to TKey
                    TKey? convertedKey = elementKey is TKey key ? key : (TKey?)Convert.ChangeType(elementKey, typeof(TKey));
                    if (convertedKey != null)
                    {
                        var elementValue = Pipeline.Map(sheet, rowIndex, valueResult, PreserveFormatting, member);
                        factory.Add(convertedKey, elementValue);
                    }
                }
            }
            finally
            {
                cellsReader.Reset();
            }

            value = factory.End();
            return true;
        }
        finally
        {
            factory.Reset();
        }
    }

    /// <summary>
    /// Sets the map that maps the value of a single cell to an object of the element type of the property
    /// or field.
    /// </summary>
    /// <param name="valueMap">The pipeline that maps the value of a single cell to an object of the element type of the property
    /// or field.</param>
    /// <returns>The map that invoked this method.</returns>
    public ManyToOneDictionaryMap<TKey, TValue> WithValueMap(Func<ValuePipeline<TValue>, ValuePipeline<TValue>> valueMap)
    {
        ThrowHelpers.ThrowIfNull(valueMap, nameof(valueMap));

        var result = valueMap(Pipeline);
        ThrowHelpers.ThrowIfNull(result, nameof(valueMap));
        Pipeline = result;

        return this;
    }

    /// <summary>
    /// Reads the dictionary from the given columns, in order.
    /// </summary>
    /// <param name="columnNames">The names of the columns.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public ManyToOneDictionaryMap<TKey, TValue> WithColumnNames(params IReadOnlyList<string> columnNames)
        => WithReaderFactory(new ColumnNamesReaderFactory(columnNames));

    /// <summary>
    /// Reads the dictionary from the given columns, in order.
    /// </summary>
    /// <param name="columnNames">The names of the columns.</param>
    /// <param name="comparison">How column names are compared.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public ManyToOneDictionaryMap<TKey, TValue> WithColumnNames(IReadOnlyList<string> columnNames, StringComparison comparison)
        => WithReaderFactory(new ColumnNamesReaderFactory(columnNames, comparison));

    /// <summary>
    /// Reads the dictionary from the given columns, in order.
    /// </summary>
    /// <param name="columnNames">The names of the columns.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public ManyToOneDictionaryMap<TKey, TValue> WithColumnNames(params IEnumerable<string> columnNames)
    {
        ThrowHelpers.ThrowIfNull(columnNames, nameof(columnNames));
        return WithColumnNames([.. columnNames]);
    }

    /// <summary>
    /// Reads the dictionary from the given columns, in order.
    /// </summary>
    /// <param name="columnNames">The names of the columns.</param>
    /// <param name="comparison">How column names are compared.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public ManyToOneDictionaryMap<TKey, TValue> WithColumnNames(IEnumerable<string> columnNames, StringComparison comparison)
    {
        ThrowHelpers.ThrowIfNull(columnNames, nameof(columnNames));
        return WithColumnNames([.. columnNames], comparison);
    }

    /// <summary>
    /// Reads the dictionary from every column whose name matches <paramref name="predicate"/>.
    /// </summary>
    /// <param name="predicate">Whether a column name matches.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public ManyToOneDictionaryMap<TKey, TValue> WithColumnsMatching(Func<string, bool> predicate)
        => WithReaderFactory(new ColumnsMatchingReaderFactory(new PredicateColumnMatcher(predicate)));

    /// <summary>
    /// Reads the dictionary from every column whose name matches <paramref name="regex"/>.
    /// </summary>
    /// <param name="regex">The regular expression column names are matched against.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public ManyToOneDictionaryMap<TKey, TValue> WithColumnsMatching(Regex regex)
        => WithReaderFactory(new ColumnsMatchingReaderFactory(new RegexColumnMatcher(regex)));

    /// <summary>
    /// Reads the dictionary from the columns at the given zero-based indices, in order.
    /// </summary>
    /// <param name="columnIndices">The zero-based indices of the columns.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public ManyToOneDictionaryMap<TKey, TValue> WithColumnIndices(params IReadOnlyList<int> columnIndices)
        => WithReaderFactory(new ColumnIndicesReaderFactory(columnIndices));

    /// <summary>
    /// Reads the dictionary from the columns at the given zero-based indices, in order.
    /// </summary>
    /// <param name="columnIndices">The zero-based indices of the columns.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public ManyToOneDictionaryMap<TKey, TValue> WithColumnIndices(params IEnumerable<int> columnIndices)
    {
        ThrowHelpers.ThrowIfNull(columnIndices, nameof(columnIndices));
        return WithColumnIndices([.. columnIndices]);
    }

    /// <summary>
    /// Reads the dictionary from the columns <paramref name="selector"/> chooses for each sheet, in order.
    /// </summary>
    /// <param name="selector">The column selector.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public ManyToOneDictionaryMap<TKey, TValue> WithColumns(IColumnSelector selector)
    {
        ThrowHelpers.ThrowIfNull(selector, nameof(selector));
        return WithReaderFactory(new ColumnSelectorReaderFactory(selector));
    }

    /// <summary>
    /// Makes the member optional: if its columns are missing, the member keeps its value instead of
    /// an error being thrown.
    /// </summary>
    /// <returns>The map on which this method was invoked.</returns>
    public ManyToOneDictionaryMap<TKey, TValue> MakeOptional()
    {
        Optional = true;
        return this;
    }

    /// <summary>
    /// Reads the cells' values as the text Excel shows, with their number formats applied.
    /// </summary>
    /// <returns>The map on which this method was invoked.</returns>
    public ManyToOneDictionaryMap<TKey, TValue> MakePreserveFormatting()
    {
        PreserveFormatting = true;
        return this;
    }

    private ManyToOneDictionaryMap<TKey, TValue> WithReaderFactory(ICellsReaderFactory readerFactory)
    {
        ReaderFactory = readerFactory;
        return this;
    }
}

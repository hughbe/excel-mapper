using ExcelMapper.Readers;
using System.Text.RegularExpressions;
using System.Collections.Concurrent;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Builders;

namespace ExcelMapper;

/// <summary>
/// A map that reads one or more values from one or more cells and maps these values to the type of the
/// property or field. This is used to map IDictionary properties and fields.
/// </summary>
/// <typeparam name="TKey">The key type of the IDictionary property or field.</typeparam>
/// <typeparam name="TValue">The value type of the IDictionary property or field.</typeparam>
public class DictionaryMap<TKey, TValue> : ICellsMap where TKey : notnull
{
    /// <summary>
    /// Constructs a map that reads values from one or more cells and maps them to the values of the
    /// dictionary assigned to the property or field.
    /// </summary>
    /// <param name="columns">The columns to read.</param>
    /// <param name="builder">Builds the dictionary.</param>
    internal DictionaryMap(Columns columns, DictionaryBuilder<TKey, TValue> builder)
    {
        _columns = columns;
        Builder = builder;
    }

    /// <summary>
    /// Gets the map that maps the value of a single cell to an object of the element type of the property
    /// or field.
    /// </summary>
    public ElementMap<TValue> Pipeline { get; private set; } = new ElementMap<TValue>();

    ValuePipeline IColumnsMap.Pipeline => Pipeline;

    private bool _frozen;

    /// <inheritdoc />
    void IMap.Freeze()
    {
        _frozen = true;
        Pipeline.Freeze();
    }

    /// <inheritdoc />
    public bool Optional
    {
        get => _optional;
        set
        {
            ThrowHelpers.ThrowIfFrozen(_frozen);
            _optional = value;
        }
    }

    private bool _optional;

    /// <inheritdoc />
    public bool PreserveFormatting
    {
        get => _preserveFormatting;
        set
        {
            ThrowHelpers.ThrowIfFrozen(_frozen);
            _preserveFormatting = value;
        }
    }

    private bool _preserveFormatting;

    private Columns _columns;

    /// <summary>
    /// Gets or sets the columns the values are read from.
    /// </summary>
    internal Columns Columns
    {
        get => _columns;
        set
        {
            ThrowHelpers.ThrowIfFrozen(_frozen);
            _columns = value;
        }
    }

    Columns IColumnsMap.Columns
    {
        get => Columns;
        set => Columns = value;
    }

    /// <summary>
    /// How the member's dictionary is built from its values.
    /// </summary>
    internal DictionaryBuilder<TKey, TValue> Builder { get; }

    /// <inheritdoc/>
    bool IMap.TryGetValue(SheetView sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, List<ExcelMappingException>? errors, [NotNullWhen(true)] out object? value)
    {
        ThrowHelpers.ThrowIfNull(sheet, nameof(sheet));
        if (sheet.Heading == null)
        {
            throw new ExcelMappingException($"The sheet \"{sheet.Name}\" does not have a heading. Use a column index map instead.");
        }

        return TryGetValue(GetCellsReader(sheet), sheet, rowIndex, reader, member, errors, out value);
    }


    ICellsReader? ICellsMap.GetCellsReader(SheetView sheet) => GetCellsReader(sheet);

    private ICellsReader? GetCellsReader(SheetView sheet)
        => sheet.GetCellsReader(_columns);

    bool ICellsMap.TryGetValue(ICellsReader? cellsReader, SheetView sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, List<ExcelMappingException>? errors, out object? value)
        => TryGetValue(cellsReader, sheet, rowIndex, reader, member, errors, out value);

    private bool TryGetValue(ICellsReader? cellsReader, SheetView sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, List<ExcelMappingException>? errors, [NotNullWhen(true)] out object? value)
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

            throw ExcelMappingException.CreateForNoSuchColumn(sheet, rowIndex, _columns, member);
        }

        var items = Builder.Begin(count);
        try
        {
            while (cellsReader.TryGetNext(out var valueResult))
            {
                // The key is the column's name.
                var elementKey = sheet.Heading.GetColumnName(valueResult.ColumnIndex);
                TKey? convertedKey = elementKey is TKey key ? key : (TKey?)Convert.ChangeType(elementKey, typeof(TKey));
                if (convertedKey != null)
                {
                    var elementValue = Pipeline.Map(sheet, rowIndex, valueResult, member, errors);
                    Builder.Add(items, convertedKey, elementValue);
                }
            }
        }
        finally
        {
            cellsReader.Reset();
        }

        value = Builder.End(items);
        return true;
    }

    /// <summary>
    /// Sets the map that maps the value of a single cell to an object of the element type of the property
    /// or field.
    /// </summary>
    /// <param name="valueMap">The pipeline that maps the value of a single cell to an object of the element type of the property
    /// or field.</param>
    /// <returns>The map that invoked this method.</returns>
    public DictionaryMap<TKey, TValue> WithValueMap(Func<ElementMap<TValue>, ElementMap<TValue>> valueMap)
    {
        ThrowHelpers.ThrowIfNull(valueMap, nameof(valueMap));

        var result = valueMap(Pipeline);
        ThrowHelpers.ThrowIfNull(result, nameof(valueMap));
        ThrowHelpers.ThrowIfFrozen(_frozen);
        Pipeline = result;

        return this;
    }

    /// <summary>
    /// Reads the dictionary from the given columns, in order.
    /// </summary>
    /// <param name="columnNames">The names of the columns.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public DictionaryMap<TKey, TValue> WithColumnNames(params IReadOnlyList<string> columnNames)
        => WithColumns(new ColumnNamesColumns(columnNames));

    /// <summary>
    /// Reads the dictionary from the given columns, in order.
    /// </summary>
    /// <param name="columnNames">The names of the columns.</param>
    /// <param name="comparison">How column names are compared.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public DictionaryMap<TKey, TValue> WithColumnNames(IReadOnlyList<string> columnNames, StringComparison comparison)
        => WithColumns(new ColumnNamesColumns(columnNames, comparison));

    /// <summary>
    /// Reads the dictionary from the given columns, in order.
    /// </summary>
    /// <param name="columnNames">The names of the columns.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public DictionaryMap<TKey, TValue> WithColumnNames(params IEnumerable<string> columnNames)
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
    public DictionaryMap<TKey, TValue> WithColumnNames(IEnumerable<string> columnNames, StringComparison comparison)
    {
        ThrowHelpers.ThrowIfNull(columnNames, nameof(columnNames));
        return WithColumnNames([.. columnNames], comparison);
    }

    /// <summary>
    /// Reads the dictionary from every column whose name matches <paramref name="predicate"/>.
    /// </summary>
    /// <param name="predicate">Whether a column name matches.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public DictionaryMap<TKey, TValue> WithColumnsMatching(Func<string, bool> predicate)
        => WithColumns(new MatchingColumns(predicate));

    /// <summary>
    /// Reads the dictionary from every column whose name matches <paramref name="regex"/>.
    /// </summary>
    /// <param name="regex">The regular expression column names are matched against.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public DictionaryMap<TKey, TValue> WithColumnsMatching(Regex regex)
        => WithColumns(new MatchingColumns(regex));

    /// <summary>
    /// Reads the dictionary from the columns at the given zero-based indices, in order.
    /// </summary>
    /// <param name="columnIndices">The zero-based indices of the columns.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public DictionaryMap<TKey, TValue> WithColumnIndices(params IReadOnlyList<int> columnIndices)
        => WithColumns(new ColumnIndicesColumns(columnIndices));

    /// <summary>
    /// Reads the dictionary from the columns at the given zero-based indices, in order.
    /// </summary>
    /// <param name="columnIndices">The zero-based indices of the columns.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public DictionaryMap<TKey, TValue> WithColumnIndices(params IEnumerable<int> columnIndices)
    {
        ThrowHelpers.ThrowIfNull(columnIndices, nameof(columnIndices));
        return WithColumnIndices([.. columnIndices]);
    }

    /// <summary>
    /// Reads the dictionary from the columns <paramref name="selector"/> chooses for each sheet, in order.
    /// </summary>
    /// <param name="selector">The column selector.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public DictionaryMap<TKey, TValue> WithColumns(IColumnSelector selector)
    {
        ThrowHelpers.ThrowIfNull(selector, nameof(selector));
        return WithColumns(new SelectorColumns(selector));
    }

    /// <summary>
    /// Makes the member optional: if its columns are missing, the member keeps its value instead of
    /// an error being thrown.
    /// </summary>
    /// <returns>The map on which this method was invoked.</returns>
    public DictionaryMap<TKey, TValue> MakeOptional()
    {
        Optional = true;
        return this;
    }

    /// <summary>
    /// Reads the cells' values as the text Excel shows, with their number formats applied.
    /// </summary>
    /// <returns>The map on which this method was invoked.</returns>
    public DictionaryMap<TKey, TValue> MakePreserveFormatting()
    {
        PreserveFormatting = true;
        return this;
    }

    private DictionaryMap<TKey, TValue> WithColumns(Columns columns)
    {
        Columns = columns;
        return this;
    }
}

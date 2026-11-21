using System.Text.RegularExpressions;
using System.Collections.Concurrent;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Readers;

namespace ExcelMapper;

/// <summary>
/// Reads multiple cells of an excel sheet and maps the value of the cell to the
/// type of the property or field.
/// </summary>
public class ManyToOneEnumerableMap<TElement> : IManyToOneMap, ICellsMap
{
    /// <summary>
    /// Constructs a map that reads one or more values from one or more cells and maps these values to one
    /// property and field of the type of the property or field.
    /// </summary>
    internal ManyToOneEnumerableMap(ICellsReaderFactory readerFactory, IEnumerableFactory<TElement> enumerableFactory)
    {
        ThrowHelpers.ThrowIfNull(readerFactory, nameof(readerFactory));
        ThrowHelpers.ThrowIfNull(enumerableFactory, nameof(enumerableFactory));

        _readerFactory = readerFactory;
        EnumerableFactory = enumerableFactory;
    }

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

    /// <inheritdoc />
    public bool Optional { get; set; }
    
    /// <inheritdoc />
    public bool PreserveFormatting { get; set; }

    /// <summary>
    /// The mapping pipeline for each element in the list.
    /// </summary>
    public ValuePipeline<TElement> Pipeline { get; private set; } = new ValuePipeline<TElement>();

    ValuePipeline IToOneMap.Pipeline => Pipeline;

    /// <summary>
    /// The factory for creating and adding elements to the list.
    /// </summary>
    internal IEnumerableFactory<TElement> EnumerableFactory { get; }

    // Factories build one value at a time, so a map shared between importers (for example
    // a class map instance registered in several importers used concurrently) must not
    // build two values at once.
    private readonly object _buildLock = new();


    /// <inheritdoc />
    bool IMap.TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, [NotNullWhen(true)] out object? value)
    {
        ThrowHelpers.ThrowIfNull(sheet, nameof(sheet));

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

    object? ICellsMap.CreateBindingFactory() => (EnumerableFactory as IBindingFactory)?.CreateForBinding();

    bool ICellsMap.TryGetValue(ICellsReader? cellsReader, object? bindingFactory, ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, out object? value)
    {
        if (bindingFactory is IEnumerableFactory<TElement> factory)
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
            return TryGetValue(cellsReader, EnumerableFactory, sheet, rowIndex, reader, member, out value);
        }
    }

    private bool TryGetValue(ICellsReader? cellsReader, IEnumerableFactory<TElement> factory, ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, [NotNullWhen(true)] out object? value)
    {
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
                while (cellsReader.TryGetNext(out var result))
                {
                    var elementValue = ValuePipeline<TElement>.Map(Pipeline, sheet, rowIndex, result, PreserveFormatting, member);
                    factory.Add(elementValue);
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
    /// <param name="elementMap">The pipeline that maps the value of a single cell to an object of the element type of the property
    /// or field.</param>
    /// <returns>The map that invoked this method.</returns>
    public ManyToOneEnumerableMap<TElement> WithElementMap(Func<ValuePipeline<TElement>, ValuePipeline<TElement>> elementMap)
    {
        ThrowHelpers.ThrowIfNull(elementMap, nameof(elementMap));

        var result = elementMap(Pipeline);
        ThrowHelpers.ThrowIfNull(result, nameof(elementMap));
        Pipeline = result;
        return this;
    }

    /// <summary>
    /// Sets the reader for multiple values to split the value of a single cell contained in the column
    /// with a given name.
    /// </summary>
    /// <param name="columnName">The name of the column containing the cell to split.</param>
    /// <param name="comparison">The string comparison to use when matching the column name.</param>
    /// <returns>The map that invoked this method.</returns>
    public ManyToOneEnumerableMap<TElement> WithColumnName(string columnName, StringComparison comparison = StringComparison.OrdinalIgnoreCase)
    {
        var columnReader = new ColumnNameReaderFactory(columnName, comparison);
        if (ReaderFactory is SplitReaderFactory splitColumnReader)
        {
            splitColumnReader.ReaderFactory = columnReader;
        }
        else
        {
            ReaderFactory = new CharSplitReaderFactory(columnReader);
        }

        return this;
    }

    /// <summary>
    /// Sets the reader for multiple values to split the value of a single cell contained in the column
    /// at the given zero-based index.
    /// </summary>
    /// <param name="columnIndex">The zero-bassed index of the column containing the cell to split.</param>
    /// <returns>The map that invoked this method.</returns>
    public ManyToOneEnumerableMap<TElement> WithColumnIndex(int columnIndex)
    {
        var factory = new ColumnIndexReaderFactory(columnIndex);
        if (ReaderFactory is SplitReaderFactory splitColumnReader)
        {
            splitColumnReader.ReaderFactory = factory;
        }
        else
        {
            ReaderFactory = new CharSplitReaderFactory(factory);
        }

        return this;
    }

    /// <summary>
    /// Sets the reader of the map to split the value of a single cell using the
    /// given separators.
    /// </summary>
    /// <param name="separators">The separators used to split the value of a single cell.</param>
    /// <returns>The map that invoked this method.</returns>
    public ManyToOneEnumerableMap<TElement> WithSeparators(params char[] separators)
    {
        SeparatorUtilities.ValidateSeparators(separators);

        if (ReaderFactory is not SplitReaderFactory splitColumnReader)
        {
            throw new ExcelMappingException("The mapping comes from multiple columns, so cannot be split.");
        }

        ReaderFactory = new CharSplitReaderFactory(splitColumnReader.ReaderFactory)
        {
            Separators = separators,
            Options = splitColumnReader.Options
        };
        return this;
    }

    /// <summary>
    /// Sets the reader of the map to split the value of a single cell using the
    /// given separators.
    /// </summary>
    /// <param name="separators">The separators used to split the value of a single cell.</param>
    /// <returns>The map that invoked this method.</returns>
    public ManyToOneEnumerableMap<TElement> WithSeparators(params IEnumerable<char> separators)
    {
        ThrowHelpers.ThrowIfNull(separators, nameof(separators));
        return WithSeparators(separators as char[] ?? separators.ToArray());
    }

    /// <summary>
    /// Sets the reader of the map to split the value of a single cell using the
    /// given separators.
    /// </summary>
    /// <param name="separators">The separators used to split the value of a single cell.</param>
    /// <returns>The map that invoked this method.</returns>
    public ManyToOneEnumerableMap<TElement> WithSeparators(params string[] separators)
    {
        SeparatorUtilities.ValidateSeparators(separators);

        if (ReaderFactory is not SplitReaderFactory splitColumnReader)
        {
            throw new ExcelMappingException("The mapping comes from multiple columns, so cannot be split.");
        }

        ReaderFactory = new StringSplitReaderFactory(splitColumnReader.ReaderFactory)
        {
            Separators = separators,
            Options = splitColumnReader.Options
        };
        return this;
    }

    /// <summary>
    /// Sets the reader of the map to split the value of a single cell using the
    /// given separators.
    /// </summary>
    /// <param name="separators">The separators used to split the value of a single cell.</param>
    /// <returns>The map that invoked this method.</returns>
    public ManyToOneEnumerableMap<TElement> WithSeparators(params IEnumerable<string> separators)
    {
        ThrowHelpers.ThrowIfNull(separators, nameof(separators));
        return WithSeparators(separators as string[] ?? separators.ToArray());
    }

    /// <summary>
    /// Reads the collection from the given columns, in order.
    /// </summary>
    /// <param name="columnNames">The names of the columns.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public ManyToOneEnumerableMap<TElement> WithColumnNames(params IReadOnlyList<string> columnNames)
        => WithReaderFactory(new ColumnNamesReaderFactory(columnNames));

    /// <summary>
    /// Reads the collection from the given columns, in order.
    /// </summary>
    /// <param name="columnNames">The names of the columns.</param>
    /// <param name="comparison">How column names are compared.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public ManyToOneEnumerableMap<TElement> WithColumnNames(IReadOnlyList<string> columnNames, StringComparison comparison)
        => WithReaderFactory(new ColumnNamesReaderFactory(columnNames, comparison));

    /// <summary>
    /// Reads the collection from the given columns, in order.
    /// </summary>
    /// <param name="columnNames">The names of the columns.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public ManyToOneEnumerableMap<TElement> WithColumnNames(params IEnumerable<string> columnNames)
    {
        ThrowHelpers.ThrowIfNull(columnNames, nameof(columnNames));
        return WithColumnNames([.. columnNames]);
    }

    /// <summary>
    /// Reads the collection from the given columns, in order.
    /// </summary>
    /// <param name="columnNames">The names of the columns.</param>
    /// <param name="comparison">How column names are compared.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public ManyToOneEnumerableMap<TElement> WithColumnNames(IEnumerable<string> columnNames, StringComparison comparison)
    {
        ThrowHelpers.ThrowIfNull(columnNames, nameof(columnNames));
        return WithColumnNames([.. columnNames], comparison);
    }

    /// <summary>
    /// Reads the collection from every column whose name matches <paramref name="predicate"/>.
    /// </summary>
    /// <param name="predicate">Whether a column name matches.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public ManyToOneEnumerableMap<TElement> WithColumnsMatching(Func<string, bool> predicate)
        => WithReaderFactory(new ColumnsMatchingReaderFactory(new PredicateColumnMatcher(predicate)));

    /// <summary>
    /// Reads the collection from every column whose name matches <paramref name="regex"/>.
    /// </summary>
    /// <param name="regex">The regular expression column names are matched against.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public ManyToOneEnumerableMap<TElement> WithColumnsMatching(Regex regex)
        => WithReaderFactory(new ColumnsMatchingReaderFactory(new RegexColumnMatcher(regex)));

    /// <summary>
    /// Reads the collection from the columns at the given zero-based indices, in order.
    /// </summary>
    /// <param name="columnIndices">The zero-based indices of the columns.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public ManyToOneEnumerableMap<TElement> WithColumnIndices(params IReadOnlyList<int> columnIndices)
        => WithReaderFactory(new ColumnIndicesReaderFactory(columnIndices));

    /// <summary>
    /// Reads the collection from the columns at the given zero-based indices, in order.
    /// </summary>
    /// <param name="columnIndices">The zero-based indices of the columns.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public ManyToOneEnumerableMap<TElement> WithColumnIndices(params IEnumerable<int> columnIndices)
    {
        ThrowHelpers.ThrowIfNull(columnIndices, nameof(columnIndices));
        return WithColumnIndices([.. columnIndices]);
    }

    /// <summary>
    /// Reads the collection from the columns <paramref name="selector"/> chooses for each sheet, in order.
    /// </summary>
    /// <param name="selector">The column selector.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public ManyToOneEnumerableMap<TElement> WithColumns(IColumnSelector selector)
    {
        ThrowHelpers.ThrowIfNull(selector, nameof(selector));
        return WithReaderFactory(new ColumnSelectorReaderFactory(selector));
    }

    /// <summary>
    /// Makes the member optional: if its columns are missing, the member keeps its value instead of
    /// an error being thrown.
    /// </summary>
    /// <returns>The map on which this method was invoked.</returns>
    public ManyToOneEnumerableMap<TElement> MakeOptional()
    {
        Optional = true;
        return this;
    }

    /// <summary>
    /// Reads the cells' values as the text Excel shows, with their number formats applied.
    /// </summary>
    /// <returns>The map on which this method was invoked.</returns>
    public ManyToOneEnumerableMap<TElement> MakePreserveFormatting()
    {
        PreserveFormatting = true;
        return this;
    }

    private ManyToOneEnumerableMap<TElement> WithReaderFactory(ICellsReaderFactory readerFactory)
    {
        ReaderFactory = readerFactory;
        return this;
    }
}

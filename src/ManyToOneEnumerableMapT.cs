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
    public ManyToOneEnumerableMap(ICellsReaderFactory readerFactory, IEnumerableFactory<TElement> enumerableFactory)
    {
        ThrowHelpers.ThrowIfNull(readerFactory, nameof(readerFactory));
        ThrowHelpers.ThrowIfNull(enumerableFactory, nameof(enumerableFactory));

        _readerFactory = readerFactory;
        EnumerableFactory = enumerableFactory;
    }

    private ICellsReaderFactory _readerFactory;

    /// <inheritdoc />
    public ICellsReaderFactory ReaderFactory
    {
        get => _readerFactory;
        set
        {
            ThrowHelpers.ThrowIfNull(value, nameof(value));
            _readerFactory = value;
            _factoryCache.Clear();
            ReaderVersion++;
        }
    }

    /// <inheritdoc />
    public bool Optional { get; set; }
    
    /// <inheritdoc />
    public bool PreserveFormatting { get; set; }

    /// <summary>
    /// The mapping pipeline for each element in the list.
    /// </summary>
    public IValuePipeline Pipeline { get; private set; } = new ValuePipeline<TElement>();

    /// <summary>
    /// The factory for creating and adding elements to the list.
    /// </summary>
    public IEnumerableFactory<TElement> EnumerableFactory { get; }

    private readonly PerSheetCache<ICellsReader> _factoryCache = new();

    // Factories build one value at a time, so a map shared between importers (for example
    // a class map instance registered in several importers used concurrently) must not
    // build two values at once.
    private readonly object _buildLock = new();


    /// <inheritdoc />
    public bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, [NotNullWhen(true)] out object? value)
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
        => _factoryCache.GetOrAdd(sheet, _readerFactory, static (s, factory) => factory.GetCellsReader(s));

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
                    var elementValue = ValuePipeline.CastValue<TElement>(ValuePipeline.GetPropertyValue(Pipeline, sheet, rowIndex, result, PreserveFormatting, member), sheet, rowIndex, result.ColumnIndex, member);
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
    public ManyToOneEnumerableMap<TElement> WithElementMap(Func<IValuePipeline<TElement>, IValuePipeline<TElement>> elementMap)
    {
        ThrowHelpers.ThrowIfNull(elementMap, nameof(elementMap));

        var result = elementMap((IValuePipeline<TElement>)Pipeline);
        ThrowHelpers.ThrowIfNull(result, nameof(elementMap));
        Pipeline = result;
        return this;
    }

    /// <summary>
    /// Sets the reader for multiple values to split the value of a single cell contained in the column
    /// with a given name.
    /// </summary>
    /// <param name="columnName">The name of the column containing the cell to split.</param>
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
}

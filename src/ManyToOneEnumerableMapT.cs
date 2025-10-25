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
public class ManyToOneEnumerableMap<TElement> : IManyToOneMap
{
    /// <summary>
    /// Constructs a map that reads one or more values from one or more cells and maps these values to one
    /// property and field of the type of the property or field.
    /// </summary>
    public ManyToOneEnumerableMap(ICellsReaderFactory readerFactory, IEnumerableFactory<TElement> enumerableFactory)
    {
        ArgumentNullException.ThrowIfNull(readerFactory);
        ArgumentNullException.ThrowIfNull(enumerableFactory);

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
            ArgumentNullException.ThrowIfNull(value);
            _readerFactory = value;
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

    private readonly ConcurrentDictionary<ExcelSheet, ICellsReader?> _factoryCache = new();

    /// <inheritdoc />
    public bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, [NotNullWhen(true)] out object? value)
    {
        ArgumentNullException.ThrowIfNull(sheet);

        var cellsReader = _factoryCache.GetOrAdd(sheet, s => _readerFactory.GetCellsReader(s));

        if (cellsReader == null || !cellsReader.TryGetValues(reader, PreserveFormatting, out var results))
        {
            if (Optional)
            {
                value = default;
                return false;
            }

            throw ExcelMappingException.CreateForNoSuchColumn(sheet, rowIndex, _readerFactory, member);
        }

        EnumerableFactory.Begin(results.Count());
        try
        {
            foreach (var result in results)
            {
                var elementValue = (TElement?)ValuePipeline.GetPropertyValue(Pipeline, sheet, rowIndex, result, PreserveFormatting, member);
                EnumerableFactory.Add(elementValue);
            }

            value = EnumerableFactory.End();
            return true;
        }
        finally
        {
            EnumerableFactory.Reset();
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
        ArgumentNullException.ThrowIfNull(elementMap);

        var result = elementMap((IValuePipeline<TElement>)Pipeline);
        ArgumentNullException.ThrowIfNull(result, nameof(elementMap));
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
        ArgumentNullException.ThrowIfNull(separators);
        return WithSeparators(separators.ToArray());
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
        ArgumentNullException.ThrowIfNull(separators);
        return WithSeparators(separators.ToArray());
    }
}

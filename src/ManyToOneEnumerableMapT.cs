using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Abstractions;
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
    public ManyToOneEnumerableMap(ICellsReaderFactory readerFactory, IValuePipeline<TElement> elementPipeline, IEnumerableFactory<TElement> enumerableFactory)
    {
        _readerFactory = readerFactory ?? throw new ArgumentNullException(nameof(readerFactory));
        ElementPipeline = elementPipeline ?? throw new ArgumentNullException(nameof(elementPipeline));
        EnumerableFactory = enumerableFactory ?? throw new ArgumentNullException(nameof(enumerableFactory));
    }

    private ICellsReaderFactory _readerFactory;

    /// <inheritdoc />
    public ICellsReaderFactory ReaderFactory
    {
        get => _readerFactory;
        set => _readerFactory = value ?? throw new ArgumentNullException(nameof(value));
    }

    /// <inheritdoc />
    public bool Optional { get; set; }
    
    /// <inheritdoc />
    public bool PreserveFormatting { get; set; }

    /// <summary>
    /// The mapping pipeline for each element in the list.
    /// </summary>
    public IValuePipeline<TElement> ElementPipeline { get; private set; }

    /// <summary>
    /// The factory for creating and adding elements to the list.
    /// </summary>
    public IEnumerableFactory<TElement> EnumerableFactory { get; }

    private readonly Dictionary<ExcelSheet, ICellsReader?> _factoryCache = [];

    public bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, [NotNullWhen(true)] out object? value)
    {
        if (sheet == null)
        {
            throw new ArgumentNullException(nameof(sheet));
        }
        if (!_factoryCache.TryGetValue(sheet, out var cellsReader))
        {
            cellsReader = _readerFactory.GetCellsReader(sheet);
            _factoryCache.Add(sheet, cellsReader);
        }

        if (cellsReader == null || !cellsReader.TryGetValues(reader, PreserveFormatting, out var results))
        {
            if (Optional)
            {
                value = default;
                return false;
            }

            throw new ExcelMappingException($"Could not read value for member \"{member?.Name}\"", sheet, rowIndex, -1);
        }

        EnumerableFactory.Begin(results.Count());
        try
        {
            foreach (var result in results)
            {
                var elementValue = (TElement?)ValuePipeline.GetPropertyValue(ElementPipeline, sheet, rowIndex, result, PreserveFormatting, member);
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
        if (elementMap == null)
        {
            throw new ArgumentNullException(nameof(elementMap));
        }

        ElementPipeline = elementMap(ElementPipeline) ?? throw new ArgumentNullException(nameof(elementMap));
        return this;
    }

    /// <summary>
    /// Sets the reader for multiple values to split the value of a single cell contained in the column
    /// with a given name.
    /// </summary>
    /// <param name="columnName">The name of the column containing the cell to split.</param>
    /// <returns>The map that invoked this method.</returns>
    public ManyToOneEnumerableMap<TElement> WithColumnName(string columnName)
    {
        var columnReader = new ColumnNameReaderFactory(columnName);
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
        if (separators is null)
        {
            throw new ArgumentNullException(nameof(separators));
        }

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
    public ManyToOneEnumerableMap<TElement> WithSeparators(IEnumerable<char> separators)
    {
        if (separators == null)
        {
            throw new ArgumentNullException(nameof(separators));
        }

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
        if (separators is null)
        {
            throw new ArgumentNullException(nameof(separators));
        }
        
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
    public ManyToOneEnumerableMap<TElement> WithSeparators(IEnumerable<string> separators)
    {
        if (separators == null)
        {
            throw new ArgumentNullException(nameof(separators));
        }

        return WithSeparators(separators.ToArray());
    }
}

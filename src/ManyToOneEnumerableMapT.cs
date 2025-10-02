using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Abstractions;
using ExcelMapper.Readers;

namespace ExcelMapper;

public delegate IEnumerable<T?> CreateElementsFactory<T>(IEnumerable<T?> elements);

/// <summary>
/// Reads multiple cells of an excel sheet and maps the value of the cell to the
/// type of the property or field.
/// </summary>
public class ManyToOneEnumerableMap<TElement> : IMap
{
    private ICellsReaderFactory _readerFactory;

    public ICellsReaderFactory ReaderFactory
    {
        get => _readerFactory;
        set => _readerFactory = value ?? throw new ArgumentNullException(nameof(value));
    }

    public bool Optional { get; set; }

    public IValuePipeline<TElement> ElementPipeline { get; private set; }

    public CreateElementsFactory<TElement> CreateElementsFactory { get; }

    /// <summary>
    /// Constructs a map that reads one or more values from one or more cells and maps these values to one
    /// property and field of the type of the property or field.
    /// </summary>
    public ManyToOneEnumerableMap(ICellsReaderFactory readerFactory, IValuePipeline<TElement> elementPipeline, CreateElementsFactory<TElement> createElementsFactory)
    {
        _readerFactory = readerFactory ?? throw new ArgumentNullException(nameof(readerFactory));
        ElementPipeline = elementPipeline ?? throw new ArgumentNullException(nameof(elementPipeline));
        CreateElementsFactory = createElementsFactory ?? throw new ArgumentNullException(nameof(createElementsFactory));
    }

    private readonly Dictionary<ExcelSheet, ICellsReader?> _factoryCache = [];

    public bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, [NotNullWhen(true)] out object? value)
    {
        if (sheet == null)
        {
            throw new ArgumentNullException(nameof(sheet));
        }
        if (!_factoryCache.TryGetValue(sheet, out ICellsReader? cellsReader))
        {
            cellsReader = _readerFactory.GetReader(sheet);
            _factoryCache.Add(sheet, cellsReader);
        }


        if (cellsReader == null || !cellsReader.TryGetValues(reader, out IEnumerable<ReadCellResult>? results))
        {
            if (Optional)
            {
                value = default;
                return false;
            }

            throw new ExcelMappingException($"Could not read value for {member?.Name}", sheet, rowIndex, -1);
        }

        var elements = new List<TElement?>();
        foreach (ReadCellResult result in results)
        {
            var elementValue = (TElement?)ValuePipeline.GetPropertyValue(ElementPipeline, sheet, rowIndex, result, member);
            elements.Add(elementValue);
        }

        value = CreateElementsFactory(elements);
        return true;
    }

    /// <summary>
    /// Makes the reader of the property map optional. For example, if the column doesn't exist
    /// or the index is invalid, an exception will not be thrown.
    /// </summary>
    /// <returns>The property map on which this method was invoked.</returns>
    public ManyToOneEnumerableMap<TElement> MakeOptional()
    {
        Optional = true;
        return this;
    }

    /// <summary>
    /// Sets the map that maps the value of a single cell to an object of the element type of the property
    /// or field.
    /// </summary>
    /// <param name="elementMap">The pipeline that maps the value of a single cell to an object of the element type of the property
    /// or field.</param>
    /// <returns>The property map that invoked this method.</returns>
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
    /// <returns>The property map that invoked this method.</returns>
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
    /// <returns>The property map that invoked this method.</returns>
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
    /// Sets the reader of the property map to split the value of a single cell using the
    /// given separators.
    /// </summary>
    /// <param name="separators">The separators used to split the value of a single cell.</param>
    /// <returns>The property map that invoked this method.</returns>
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
    /// Sets the reader of the property map to split the value of a single cell using the
    /// given separators.
    /// </summary>
    /// <param name="separators">The separators used to split the value of a single cell.</param>
    /// <returns>The property map that invoked this method.</returns>
    public ManyToOneEnumerableMap<TElement> WithSeparators(IEnumerable<char> separators)
    {
        if (separators == null)
        {
            throw new ArgumentNullException(nameof(separators));
        }

        return WithSeparators(separators.ToArray());
    }

    /// <summary>
    /// Sets the reader of the property map to split the value of a single cell using the
    /// given separators.
    /// </summary>
    /// <param name="separators">The separators used to split the value of a single cell.</param>
    /// <returns>The property map that invoked this method.</returns>
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
    /// Sets the reader of the property map to split the value of a single cell using the
    /// given separators.
    /// </summary>
    /// <param name="separators">The separators used to split the value of a single cell.</param>
    /// <returns>The property map that invoked this method.</returns>
    public ManyToOneEnumerableMap<TElement> WithSeparators(IEnumerable<string> separators)
    {
        if (separators == null)
        {
            throw new ArgumentNullException(nameof(separators));
        }

        return WithSeparators(separators.ToArray());
    }

    /// <summary>
    /// Sets the reader of the property map to read the values of one or more cells contained
    /// in the columns with the given names.
    /// </summary>
    /// <param name="columnNames">The name of each column to read.</param>
    /// <returns>The property map that invoked this method.</returns>
    public ManyToOneEnumerableMap<TElement> WithColumnNames(params string[] columnNames)
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
    public ManyToOneEnumerableMap<TElement> WithColumnNames(IEnumerable<string> columnNames)
    {
        if (columnNames == null)
        {
            throw new ArgumentNullException(nameof(columnNames));
        }

        return WithColumnNames([.. columnNames]);
    }

    /// <summary>
    /// Sets the reader of the property map to read the values of one or more cells contained
    /// in the columns with the given zero-based indices.
    /// </summary>
    /// <param name="columnIndices">The zero-based index of each column to read.</param>
    /// <returns>The property map that invoked this method.</returns>
    public ManyToOneEnumerableMap<TElement> WithColumnIndices(params int[] columnIndices)
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
    public ManyToOneEnumerableMap<TElement> WithColumnIndices(IEnumerable<int> columnIndices)
    {
        if (columnIndices == null)
        {
            throw new ArgumentNullException(nameof(columnIndices));
        }
        
        return WithColumnIndices([.. columnIndices]);
    }
}

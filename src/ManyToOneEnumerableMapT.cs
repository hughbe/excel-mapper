#if MULTIPLE
using System.Linq;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Abstractions;
using ExcelMapper.Readers;

namespace ExcelMapper;
 
public delegate IEnumerable<T> CreateElementsFactory<T>(IEnumerable<T> elements);

/// <summary>
/// Reads multiple cells of an excel sheet and maps the value of the cell to the
/// type of the property or field.
/// </summary>
public class ManyToOneEnumerableMap<TElement> : IMap
{
    public bool Optional { get; set; }

    public IEnumerable<IMapper<TElement>> Mappers { get; }

    public CreateElementsFactory<TElement> CreateElementsFactory { get; }

    /// <summary>
    /// Constructs a map that reads one or more values from one or more cells and maps these values to one
    /// property and field of the type of the property or field.
    /// </summary>
    public ManyToOneEnumerableMap(IEnumerable<IMapper<TElement>> mappers, CreateElementsFactory<TElement> createElementsFactory)
    {
        Mappers = mappers ?? throw new ArgumentNullException(nameof(mappers));
        CreateElementsFactory = createElementsFactory ?? throw new ArgumentNullException(nameof(createElementsFactory));
    }

    public bool TryGetValue(ExcelRow row, IExcelDataReader reader, MemberInfo member, out object value)
    {
        var elements = new List<TElement>();
        foreach (IMapper<TElement> mapper in Mappers)
        {
            if (mapper.TryGetValue(row, reader, member, out object element))
            {
                elements.Add((TElement)element);
            }
            else if (!Optional)
            {
                value = null;
                return false;
            }
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

        elementMapper = elementMap(elementMapper) ?? throw new ArgumentNullException(nameof(elementMap));
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
        var columnReader = new ColumnNameValueReader(columnName);
        if (CellValuesReader is SplitCellValueReader splitColumnReader)
        {
            splitColumnReader.CellReader = columnReader;
        }
        else
        {
            CellValuesReader = new CharSplitCellValueReader(columnReader);
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
        var reader = new ColumnIndexValueReader(columnIndex);
        if (CellValuesReader is SplitCellValueReader splitColumnReader)
        {
            splitColumnReader.CellReader = reader;
        }
        else
        {
            CellValuesReader = new CharSplitCellValueReader(reader);
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
        if (!(CellValuesReader is SplitCellValueReader splitColumnReader))
        {
            throw new ExcelMappingException("The mapping comes from multiple columns, so cannot be split.");
        }

        CellValuesReader = new CharSplitCellValueReader(splitColumnReader.CellReader)
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
        return WithSeparators(separators?.ToArray());
    }

    /// <summary>
    /// Sets the reader of the property map to split the value of a single cell using the
    /// given separators.
    /// </summary>
    /// <param name="separators">The separators used to split the value of a single cell.</param>
    /// <returns>The property map that invoked this method.</returns>
    public ManyToOneEnumerableMap<TElement> WithSeparators(params string[] separators)
    {
        if (!(CellValuesReader is SplitCellValueReader splitColumnReader))
        {
            throw new ExcelMappingException("The mapping comes from multiple columns, so cannot be split.");
        }

        CellValuesReader = new StringSplitCellValueReader(splitColumnReader.CellReader)
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
        return WithSeparators(separators?.ToArray());
    }

    /// <summary>
    /// Sets the reader of the property map to read the values of one or more cells contained
    /// in the columns with the given names.
    /// </summary>
    /// <param name="columnNames">The name of each column to read.</param>
    /// <returns>The property map that invoked this method.</returns>
    public ManyToOneEnumerableMap<TElement> WithColumnNames(params string[] columnNames)
    {
        CellValuesReader = new MultipleColumnNamesValueReader(columnNames);
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
        return WithColumnNames(columnNames?.ToArray());
    }

    /// <summary>
    /// Sets the reader of the property map to read the values of one or more cells contained
    /// in the columns with the given zero-based indices.
    /// </summary>
    /// <param name="columnIndices">The zero-based index of each column to read.</param>
    /// <returns>The property map that invoked this method.</returns>
    public ManyToOneEnumerableMap<TElement> WithColumnIndices(params int[] columnIndices)
    {
        CellValuesReader = new MultipleColumnIndicesValueReader(columnIndices);
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
        return WithColumnIndices(columnIndices?.ToArray());
    }
}
#endif

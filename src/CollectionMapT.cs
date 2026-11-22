using System.Text.RegularExpressions;
using System.Collections.Concurrent;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Reflection;
using ExcelDataReader;
using ExcelMapper.Readers;
using ExcelMapper.Factories;

namespace ExcelMapper;

/// <summary>
/// Reads multiple cells of an excel sheet and maps the value of the cell to the
/// type of the property or field.
/// </summary>
public class CollectionMap<TElement> : ICellsMap
{
    /// <summary>
    /// Constructs a map that reads one or more values from one or more cells and maps these values to one
    /// property and field of the type of the property or field.
    /// </summary>
    internal CollectionMap(Columns columns, CollectionBuilder<TElement> builder)
    {
        _columns = columns;
        Builder = builder;
    }

    private Columns _columns;

    /// <summary>
    /// Gets or sets the columns the elements are read from: several columns, or one cell split
    /// into values (<see cref="SplitColumns"/>).
    /// </summary>
    internal Columns Columns
    {
        get => _columns;
        set
        {
            _columns = value;
            ReaderVersion++;
        }
    }

    Columns IColumnsMap.Columns
    {
        get => Columns;
        set => Columns = value;
    }

    /// <inheritdoc />
    public bool Optional { get; set; }
    
    /// <inheritdoc />
    public bool PreserveFormatting { get; set; }

    /// <summary>
    /// The mapping pipeline for each element in the list.
    /// </summary>
    public ElementMap<TElement> Pipeline { get; private set; } = new ElementMap<TElement>();

    ValuePipeline IColumnsMap.Pipeline => Pipeline;

    /// <summary>
    /// How the member's collection is built from its elements.
    /// </summary>
    internal CollectionBuilder<TElement> Builder { get; }

    /// <inheritdoc />
    bool IMap.TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, List<ExcelMappingException>? errors, [NotNullWhen(true)] out object? value)
    {
        ThrowHelpers.ThrowIfNull(sheet, nameof(sheet));

        return TryGetValue(GetCellsReader(sheet), sheet, rowIndex, reader, member, errors, out value);
    }

    /// <summary>
    /// Incremented when <see cref="Columns"/> changes, so that readers resolved for a sheet
    /// (see <see cref="BoundCellsMember"/>) are resolved again.
    /// </summary>
    int ICellsMap.ReaderVersion => ReaderVersion;

    internal int ReaderVersion { get; private set; }

    ICellsReader? ICellsMap.GetCellsReader(ExcelSheet sheet) => GetCellsReader(sheet);

    private ICellsReader? GetCellsReader(ExcelSheet sheet)
        => sheet.GetCellsReader(_columns);

    bool ICellsMap.TryGetValue(ICellsReader? cellsReader, ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, List<ExcelMappingException>? errors, out object? value)
        => TryGetValue(cellsReader, sheet, rowIndex, reader, member, errors, out value);

    private bool TryGetValue(ICellsReader? cellsReader, ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, List<ExcelMappingException>? errors, [NotNullWhen(true)] out object? value)
    {
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
            var index = 0;
            while (cellsReader.TryGetNext(out var result))
            {
                var elementValue = Pipeline.Map(sheet, rowIndex, result, member, errors);
                Builder.Add(items, index++, elementValue);
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
    /// <param name="elementMap">The pipeline that maps the value of a single cell to an object of the element type of the property
    /// or field.</param>
    /// <returns>The map that invoked this method.</returns>
    public CollectionMap<TElement> WithElementMap(Func<ElementMap<TElement>, ElementMap<TElement>> elementMap)
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
    public CollectionMap<TElement> WithColumnName(string columnName, StringComparison comparison = StringComparison.OrdinalIgnoreCase)
        => WithSplitCell(new ColumnNameColumns(columnName, comparison));

    /// <summary>
    /// Sets the reader for multiple values to split the value of a single cell contained in the column
    /// at the given zero-based index.
    /// </summary>
    /// <param name="columnIndex">The zero-bassed index of the column containing the cell to split.</param>
    /// <returns>The map that invoked this method.</returns>
    public CollectionMap<TElement> WithColumnIndex(int columnIndex)
        => WithSplitCell(new ColumnIndexColumns(columnIndex));

    // Splits the cell in the given column, with the separators set before, if any.
    private CollectionMap<TElement> WithSplitCell(Columns cell)
    {
        Columns = new SplitColumns(cell, Columns is SplitColumns split ? split.Splitter : new CharSplitter());
        return this;
    }

    /// <summary>
    /// Sets the reader of the map to split the value of a single cell using the
    /// given separators.
    /// </summary>
    /// <param name="separators">The separators used to split the value of a single cell.</param>
    /// <returns>The map that invoked this method.</returns>
    public CollectionMap<TElement> WithSeparators(params char[] separators)
    {
        SeparatorUtilities.ValidateSeparators(separators);
        return WithSplitter(new CharSplitter { Separators = separators });
    }

    /// <summary>
    /// Sets the reader of the map to split the value of a single cell using the
    /// given separators.
    /// </summary>
    /// <param name="separators">The separators used to split the value of a single cell.</param>
    /// <returns>The map that invoked this method.</returns>
    public CollectionMap<TElement> WithSeparators(params IEnumerable<char> separators)
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
    public CollectionMap<TElement> WithSeparators(params string[] separators)
    {
        SeparatorUtilities.ValidateSeparators(separators);
        return WithSplitter(new StringSplitter { Separators = separators });
    }

    // Splits the cell with the given separators, keeping the split options.
    private CollectionMap<TElement> WithSplitter(CellSplitter splitter)
    {
        if (Columns is not SplitColumns split)
        {
            throw new ExcelMappingException("The mapping comes from multiple columns, so cannot be split.");
        }

        splitter.Options = split.Splitter.Options;
        Columns = new SplitColumns(split.Cell, splitter);
        return this;
    }

    /// <summary>
    /// Sets the reader of the map to split the value of a single cell using the
    /// given separators.
    /// </summary>
    /// <param name="separators">The separators used to split the value of a single cell.</param>
    /// <returns>The map that invoked this method.</returns>
    public CollectionMap<TElement> WithSeparators(params IEnumerable<string> separators)
    {
        ThrowHelpers.ThrowIfNull(separators, nameof(separators));
        return WithSeparators(separators as string[] ?? separators.ToArray());
    }

    /// <summary>
    /// Reads the collection from the given columns, in order.
    /// </summary>
    /// <param name="columnNames">The names of the columns.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public CollectionMap<TElement> WithColumnNames(params IReadOnlyList<string> columnNames)
        => WithColumns(new ColumnNamesColumns(columnNames));

    /// <summary>
    /// Reads the collection from the given columns, in order.
    /// </summary>
    /// <param name="columnNames">The names of the columns.</param>
    /// <param name="comparison">How column names are compared.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public CollectionMap<TElement> WithColumnNames(IReadOnlyList<string> columnNames, StringComparison comparison)
        => WithColumns(new ColumnNamesColumns(columnNames, comparison));

    /// <summary>
    /// Reads the collection from the given columns, in order.
    /// </summary>
    /// <param name="columnNames">The names of the columns.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public CollectionMap<TElement> WithColumnNames(params IEnumerable<string> columnNames)
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
    public CollectionMap<TElement> WithColumnNames(IEnumerable<string> columnNames, StringComparison comparison)
    {
        ThrowHelpers.ThrowIfNull(columnNames, nameof(columnNames));
        return WithColumnNames([.. columnNames], comparison);
    }

    /// <summary>
    /// Reads the collection from every column whose name matches <paramref name="predicate"/>.
    /// </summary>
    /// <param name="predicate">Whether a column name matches.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public CollectionMap<TElement> WithColumnsMatching(Func<string, bool> predicate)
        => WithColumns(new MatchingColumns(predicate));

    /// <summary>
    /// Reads the collection from every column whose name matches <paramref name="regex"/>.
    /// </summary>
    /// <param name="regex">The regular expression column names are matched against.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public CollectionMap<TElement> WithColumnsMatching(Regex regex)
        => WithColumns(new MatchingColumns(regex));

    /// <summary>
    /// Reads the collection from the columns at the given zero-based indices, in order.
    /// </summary>
    /// <param name="columnIndices">The zero-based indices of the columns.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public CollectionMap<TElement> WithColumnIndices(params IReadOnlyList<int> columnIndices)
        => WithColumns(new ColumnIndicesColumns(columnIndices));

    /// <summary>
    /// Reads the collection from the columns at the given zero-based indices, in order.
    /// </summary>
    /// <param name="columnIndices">The zero-based indices of the columns.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public CollectionMap<TElement> WithColumnIndices(params IEnumerable<int> columnIndices)
    {
        ThrowHelpers.ThrowIfNull(columnIndices, nameof(columnIndices));
        return WithColumnIndices([.. columnIndices]);
    }

    /// <summary>
    /// Reads the collection from the columns <paramref name="selector"/> chooses for each sheet, in order.
    /// </summary>
    /// <param name="selector">The column selector.</param>
    /// <returns>The map on which this method was invoked.</returns>
    public CollectionMap<TElement> WithColumns(IColumnSelector selector)
    {
        ThrowHelpers.ThrowIfNull(selector, nameof(selector));
        return WithColumns(new SelectorColumns(selector));
    }

    /// <summary>
    /// Makes the member optional: if its columns are missing, the member keeps its value instead of
    /// an error being thrown.
    /// </summary>
    /// <returns>The map on which this method was invoked.</returns>
    public CollectionMap<TElement> MakeOptional()
    {
        Optional = true;
        return this;
    }

    /// <summary>
    /// Reads the cells' values as the text Excel shows, with their number formats applied.
    /// </summary>
    /// <returns>The map on which this method was invoked.</returns>
    public CollectionMap<TElement> MakePreserveFormatting()
    {
        PreserveFormatting = true;
        return this;
    }

    private CollectionMap<TElement> WithColumns(Columns columns)
    {
        Columns = columns;
        return this;
    }
}

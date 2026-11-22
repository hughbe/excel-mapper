using ExcelMapper.Readers;

namespace ExcelMapper.Utilities;

/// <summary>
/// What's been worked out about a sheet's columns, so it's done once per sheet rather than for
/// every row: the columns members read, the readers of collections' cells, and class maps
/// bound to it.
/// </summary>
internal sealed class SheetBindings(SheetView sheet)
{
    // Keyed by the columns instance (the kinds of columns don't override Equals): a map whose
    // columns are replaced finds them again.
    private Dictionary<Columns, int>? _columns;
    private Dictionary<Columns, ICellsReader?>? _cellsReaders;

    // Extra columns depend on the root map as well as the columns, as a nested class map can be
    // shared between root maps (with WithClassMap).
    private Dictionary<(Columns, IMap?), ICellsReader?>? _extraColumnsReaders;

    /// <summary>
    /// Finds the column a single value is read from, the first time it's asked for.
    /// </summary>
    public int FindColumn(Columns columns)
    {
        _columns ??= [];
        if (!_columns.TryGetValue(columns, out var column))
        {
            column = columns.FindColumn(sheet);
            _columns[columns] = column;
        }

        return column;
    }

    /// <summary>
    /// Gets the reader of the cells a collection or dictionary is read from, creating it the
    /// first time it's asked for. Null if the sheet doesn't have the columns.
    /// </summary>
    public ICellsReader? GetCellsReader(Columns columns)
    {
        if (columns is ExtraColumns)
        {
            _extraColumnsReaders ??= [];
            var key = (columns, sheet.RootMap);
            if (!_extraColumnsReaders.TryGetValue(key, out var extraColumnsReader))
            {
                extraColumnsReader = CreateCellsReader(columns);
                _extraColumnsReaders[key] = extraColumnsReader;
            }

            return extraColumnsReader;
        }

        _cellsReaders ??= [];
        if (!_cellsReaders.TryGetValue(columns, out var reader))
        {
            reader = CreateCellsReader(columns);
            _cellsReaders[columns] = reader;
        }

        return reader;
    }

    private ICellsReader? CreateCellsReader(Columns columns)
    {
        if (columns is SplitColumns split)
        {
            var column = split.FindColumn(sheet);
            return column == -1 ? null : new SplitCellsReader(column, split.Splitter);
        }

        return columns.FindColumns(sheet) is { } found ? new ColumnIndicesReader(found) : null;
    }

    // Class maps bound to the sheet. Rows are usually mapped with one class map at a time, so the
    // last one is kept separately. Class maps are frozen before rows are read, so their members
    // never need binding again.
    private Dictionary<ExcelClassMap, BoundMember[]>? _boundClassMaps;
    private ExcelClassMap? _lastBoundClassMap;
    private BoundMember[]? _lastBound;

    /// <summary>
    /// Gets the members of a class map bound to the sheet, binding them the first time.
    /// </summary>
    public BoundMember[] GetBoundMembers(ExcelClassMap classMap)
    {
        if (ReferenceEquals(_lastBoundClassMap, classMap))
        {
            return _lastBound!;
        }

        _boundClassMaps ??= [];
        if (!_boundClassMaps.TryGetValue(classMap, out var bound))
        {
            bound = Bind(classMap);
            _boundClassMaps[classMap] = bound;
        }

        _lastBoundClassMap = classMap;
        _lastBound = bound;
        return bound;
    }

    private BoundMember[] Bind(ExcelClassMap classMap)
    {
        var properties = classMap.Properties;
        var members = new BoundMember[properties.Count];
        for (int i = 0; i < members.Length; i++)
        {
            var property = properties[i];
            members[i] = property.Map switch
            {
                IValueMap bindableMap => bindableMap.Bind(property, sheet),
                ICellsMap cellsMap => new BoundCellsMember(cellsMap, property),
                _ => new BoundMapMember(property),
            };
        }

        return members;
    }
}

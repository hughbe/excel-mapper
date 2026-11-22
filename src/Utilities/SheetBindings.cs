using ExcelMapper.Readers;

namespace ExcelMapper.Utilities;

/// <summary>
/// What's been worked out about a sheet's columns, so it's done once per sheet rather than for
/// every row: the readers that reader factories create for it, and class maps bound to it.
/// </summary>
internal sealed class SheetBindings(ExcelSheet sheet)
{
    // Keyed by the factory instance (the factories are internal and don't override Equals): a map
    // whose factory is replaced gets a new reader.
    private Dictionary<ICellReaderFactory, ICellReader?>? _cellReaders;
    private Dictionary<ICellsReaderFactory, ICellsReader?>? _cellsReaders;

    // Extra columns depend on the root map as well as the factory, as a nested class map can be
    // shared between root maps (with WithClassMap).
    private Dictionary<(ExtraColumnsReaderFactory, IMap?), ICellsReader?>? _extraColumnsReaders;

    /// <summary>
    /// Gets the reader a factory creates for the sheet, creating it the first time.
    /// </summary>
    public ICellReader? GetCellReader(ICellReaderFactory factory)
    {
        _cellReaders ??= [];
        if (!_cellReaders.TryGetValue(factory, out var reader))
        {
            reader = factory.GetCellReader(sheet);
            _cellReaders[factory] = reader;
        }

        return reader;
    }

    /// <summary>
    /// Gets the cells reader a factory creates for the sheet, creating it the first time.
    /// </summary>
    public ICellsReader? GetCellsReader(ICellsReaderFactory factory)
    {
        if (factory is ExtraColumnsReaderFactory extraColumnsFactory)
        {
            _extraColumnsReaders ??= [];
            var key = (extraColumnsFactory, sheet.RootMap);
            if (!_extraColumnsReaders.TryGetValue(key, out var extraColumnsReader))
            {
                extraColumnsReader = extraColumnsFactory.GetCellsReader(sheet);
                _extraColumnsReaders[key] = extraColumnsReader;
            }

            return extraColumnsReader;
        }

        _cellsReaders ??= [];
        if (!_cellsReaders.TryGetValue(factory, out var reader))
        {
            reader = factory.GetCellsReader(sheet);
            _cellsReaders[factory] = reader;
        }

        return reader;
    }

    private sealed class BoundClassMap(int version, BoundMember[] members)
    {
        public int Version { get; } = version;

        public BoundMember[] Members { get; } = members;
    }

    // Class maps bound to the sheet. Rows are usually mapped with one class map at a time, so the
    // last one is kept separately.
    private Dictionary<ExcelClassMap, BoundClassMap>? _boundClassMaps;
    private ExcelClassMap? _lastBoundClassMap;
    private BoundClassMap? _lastBound;

    /// <summary>
    /// Gets the members of a class map bound to the sheet, binding them again if the class map has
    /// changed.
    /// </summary>
    public BoundMember[] GetBoundMembers(ExcelClassMap classMap)
    {
        var bound = ReferenceEquals(_lastBoundClassMap, classMap) ? _lastBound : null;
        if (bound is null)
        {
            _boundClassMaps ??= [];
            _boundClassMaps.TryGetValue(classMap, out bound);
        }

        var version = classMap.PropertiesVersion;
        if (bound is null || bound.Version != version)
        {
            bound = new BoundClassMap(version, Bind(classMap));
            _boundClassMaps![classMap] = bound;
        }

        _lastBoundClassMap = classMap;
        _lastBound = bound;
        return bound.Members;
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
                IBindableMap bindableMap => bindableMap.Bind(property, sheet),
                ICellsMap cellsMap => new BoundCellsMember(cellsMap, property),
                _ => new BoundMapMember(property),
            };
        }

        return members;
    }
}

using System.Diagnostics.CodeAnalysis;
using ExcelDataReader;

namespace ExcelMapper.Readers;

/// <summary>
/// Reads every column in the heading that no other member of the map being read reads. Used by
/// <see cref="ExcelExtraColumnsAttribute"/> and
/// <see cref="ExcelClassMap{T}.MapExtraColumns{TValue}(System.Linq.Expressions.Expression{Func{T, IEnumerable{KeyValuePair{string, TValue}}}})"/>.
/// </summary>
/// <remarks>
/// The columns are worked out once per sheet from the members of the sheet's root map (see
/// <see cref="ExcelSheet.RootMap"/>), including members of nested objects at any depth, so a
/// member of a nested object doesn't get the columns its parent's members read. Columns with an
/// empty name are skipped.
/// </remarks>
internal sealed class ExtraColumnsReaderFactory : ICellsReaderFactory
{
    /// <summary>
    /// Initializes a new instance of <see cref="ExtraColumnsReaderFactory"/>.
    /// </summary>
    /// <param name="classMap">The class map that has the extra columns member. Its members' columns
    /// are excluded if the sheet isn't being read with a root map.</param>
    /// <exception cref="ArgumentNullException"><paramref name="classMap"/> is <see langword="null"/>.</exception>
    public ExtraColumnsReaderFactory(ExcelClassMap classMap)
    {
        ThrowHelpers.ThrowIfNull(classMap, nameof(classMap));
        ClassMap = classMap;
    }

    /// <summary>
    /// Gets the class map that has the extra columns member.
    /// </summary>
    public ExcelClassMap ClassMap { get; }

    /// <inheritdoc/>
    public ICellsReader? GetCellsReader(ExcelSheet sheet)
    {
        ThrowHelpers.ThrowIfNull(sheet, nameof(sheet));
        if (sheet.Heading == null)
        {
            throw new ExcelMappingException($"The sheet \"{sheet.Name}\" does not have a heading. Extra columns can only be read from sheets with a heading.");
        }

        var usedColumns = new HashSet<int>();
        AddUsedColumns(sheet.RootMap ?? ClassMap, sheet, usedColumns);

        var heading = sheet.Heading;
        var columnNames = heading.ColumnNames;
        var indices = new List<int>();
        for (int i = 0; i < columnNames.Count; i++)
        {
            var columnIndex = i + heading.ColumnOffset;
            if (!string.IsNullOrWhiteSpace(columnNames[i]) && !usedColumns.Contains(columnIndex))
            {
                indices.Add(columnIndex);
            }
        }

        return indices.Count == 0 ? EmptyCellsReader.Instance : new ColumnIndicesReader(indices.ToArray());
    }

    private static void AddUsedColumns(IMap map, ExcelSheet sheet, HashSet<int> usedColumns)
    {
        switch (map)
        {
            case ExcelClassMap classMap:
                if (classMap._valueMap is not null)
                {
                    AddUsedColumns(classMap._valueMap, sheet, usedColumns);
                }

                var properties = classMap.Properties;
                for (int i = 0; i < properties.Count; i++)
                {
                    AddUsedColumns(properties[i].Map, sheet, usedColumns);
                }

                break;

            case IValueMap valueMap:
                AddUsedColumns(valueMap.ReaderFactory, sheet, usedColumns);
                break;

            case IManyToOneMap { ReaderFactory: ExtraColumnsReaderFactory }:
                // Another extra columns member: it doesn't claim any columns.
                break;

            case IManyToOneMap cellsMap:
                AddUsedColumns(cellsMap.ReaderFactory, sheet, usedColumns);
                break;

            case IEnumerableIndexerMap enumerableIndexerMap:
                foreach (var value in enumerableIndexerMap.Values.Values)
                {
                    AddUsedColumns(value, sheet, usedColumns);
                }

                break;

            case IMultidimensionalIndexerMap multidimensionalIndexerMap:
                foreach (var value in multidimensionalIndexerMap.Values.Values)
                {
                    AddUsedColumns(value, sheet, usedColumns);
                }

                break;

            case IDictionaryIndexerMap dictionaryIndexerMap:
                foreach (var value in dictionaryIndexerMap.Values.Values)
                {
                    AddUsedColumns(value, sheet, usedColumns);
                }

                break;
        }
    }

    private static void AddUsedColumns(ICellReaderFactory readerFactory, ExcelSheet sheet, HashSet<int> usedColumns)
    {
        try
        {
            AddUsedColumns(readerFactory.GetCellReader(sheet), usedColumns);
        }
        catch (ExcelMappingException)
        {
            // The member's column can't be found (for example an optional column that is
            // missing), so it doesn't read any of the sheet's columns.
        }
    }

    private static void AddUsedColumns(ICellsReaderFactory readerFactory, ExcelSheet sheet, HashSet<int> usedColumns)
    {
        try
        {
            if (readerFactory is SplitReaderFactory splitReaderFactory)
            {
                // Splitting reads values from a single column.
                AddUsedColumns(splitReaderFactory.ReaderFactory.GetCellReader(sheet), usedColumns);
                return;
            }

            switch (readerFactory.GetCellsReader(sheet))
            {
                case ColumnIndicesReader columnIndicesReader:
                    foreach (var columnIndex in columnIndicesReader.ColumnIndices)
                    {
                        usedColumns.Add(columnIndex);
                    }

                    break;
            }
        }
        catch (ExcelMappingException)
        {
            // The member's columns can't be found, so it doesn't read any of the sheet's columns.
        }
    }

    private static void AddUsedColumns(ICellReader? reader, HashSet<int> usedColumns)
    {
        if (reader is ColumnIndexReader columnIndexReader)
        {
            usedColumns.Add(columnIndexReader.ColumnIndex);
        }
    }

    private sealed class EmptyCellsReader : ICellsReader
    {
        public static EmptyCellsReader Instance { get; } = new();

        public bool Start(IExcelDataReader reader, bool preserveFormatting, out int count)
        {
            count = 0;
            return true;
        }

        public bool TryGetNext([NotNullWhen(true)] out ReadCellResult result)
        {
            result = default;
            return false;
        }

        public void Reset()
        {
        }
    }
}

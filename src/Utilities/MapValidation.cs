using ExcelMapper.Readers;

namespace ExcelMapper.Utilities;

/// <summary>
/// Finds the columns a map needs that a sheet doesn't have.
/// </summary>
internal static class MapValidation
{
    // Adds a problem for each column the map needs that the sheet doesn't have.
    // reportAmbiguous: whether a column name that more than one column matches is a problem. It
    // is when validating a heading; when searching for one, the row is still the heading, and
    // reading it reports which column to choose.
    public static void FindMissingColumns(IMap map, SheetView sheet, string? path, List<string> problems, bool reportAmbiguous = true)
        => MapTree.VisitCellMaps(map, path, (cellMap, memberPath) =>
        {
            if (!cellMap.Settings.Optional)
            {
                Check(cellMap.Settings.Columns, many: cellMap is ICellsMap, sheet, memberPath, problems, reportAmbiguous);
            }
        });

    public static void CountNamedColumns(IMap map, SheetView sheet, ref int lookedUp, ref int found)
    {
        var lookedUpCount = 0;
        var foundCount = 0;
        MapTree.VisitCellMaps(map, null, (cellMap, _) =>
        {
            var columns = cellMap.Settings.Columns;
            if (columns.LooksUpNames)
            {
                lookedUpCount++;
                if (HasColumns(columns, many: cellMap is ICellsMap, sheet))
                {
                    foundCount++;
                }
            }
        });

        lookedUp += lookedUpCount;
        found += foundCount;
    }

    // A collection split from one cell reads one column.
    private static bool Found(Columns columns, bool many, SheetView sheet)
        => many && columns is not SplitColumns ? columns.FindColumns(sheet) is not null : columns.FindColumn(sheet) != -1;

    private static bool HasColumns(Columns columns, bool many, SheetView sheet)
    {
        try
        {
            return Found(columns, many, sheet);
        }
        catch (ExcelMappingException exception)
        {
            // A column whose name appears twice is there: reading it reports which to choose.
            return exception.IsAmbiguousColumn;
        }
    }

    private static void Check(Columns columns, bool many, SheetView sheet, string? path, List<string> problems, bool reportAmbiguous)
    {
        string? error = null;
        try
        {
            if (Found(columns, many, sheet))
            {
                return;
            }
        }
        catch (ExcelMappingException exception)
        {
            if (exception.IsAmbiguousColumn && !reportAmbiguous)
            {
                return;
            }

            // For example, a column looked up by name in a sheet without a heading.
            error = exception.Message;
        }

        var member = path is null ? string.Empty : $" (member {path})";
        problems.Add($"{error ?? columns.Describe(sheet)}{member}");
    }
}

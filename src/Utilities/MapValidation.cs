using ExcelMapper.Readers;

namespace ExcelMapper.Utilities;

/// <summary>
/// Finds the columns a map needs that a sheet doesn't have.
/// </summary>
internal static class MapValidation
{
    public static void FindMissingColumns(IMap map, ExcelSheet sheet, string? path, List<string> problems)
        => MapTree.VisitCellMaps(map, path, (cellMap, memberPath) =>
        {
            if (!cellMap.Optional)
            {
                Check(cellMap.Columns, many: cellMap is ICellsMap, sheet, memberPath, problems);
            }
        });

    public static void CountNamedColumns(IMap map, ExcelSheet sheet, ref int lookedUp, ref int found)
    {
        var lookedUpCount = 0;
        var foundCount = 0;
        MapTree.VisitCellMaps(map, null, (cellMap, _) =>
        {
            var columns = cellMap.Columns;
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
    private static bool Found(Columns columns, bool many, ExcelSheet sheet)
        => many && columns is not SplitColumns ? columns.FindColumns(sheet) is not null : columns.FindColumn(sheet) != -1;

    private static bool HasColumns(Columns columns, bool many, ExcelSheet sheet)
    {
        try
        {
            return Found(columns, many, sheet);
        }
        catch (ExcelMappingException)
        {
            return false;
        }
    }

    private static void Check(Columns columns, bool many, ExcelSheet sheet, string? path, List<string> problems)
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
            // For example, a column looked up by name in a sheet without a heading.
            error = exception.Message;
        }

        var member = path is null ? string.Empty : $" (member {path})";
        problems.Add($"{error ?? columns.Describe(sheet)}{member}");
    }
}

using System.Linq;

namespace ExcelMapper.Readers;

/// <summary>
/// The columns with a name that no other member of a class map reads: see
/// <see cref="ExcelExtraColumnsAttribute"/>. When rows are read, the class map is the one rows
/// are mapped with (see <see cref="SheetView.RootMap"/>), as a nested class map can be shared.
/// </summary>
internal sealed class ExtraColumns : Columns
{
    public ExtraColumns(ExcelClassMap classMap)
    {
        ThrowHelpers.ThrowIfNull(classMap, nameof(classMap));
        ClassMap = classMap;
    }

    public ExcelClassMap ClassMap { get; }

    public override int FindColumn(SheetView sheet) => throw new NotSupportedException();

    public override int[] FindColumns(SheetView sheet)
    {
        var heading = sheet.Heading ?? throw new ExcelMappingException($"The sheet \"{sheet.Name}\" does not have a heading. Extra columns can only be read from sheets with a heading.");

        var usedColumns = new HashSet<int>();
        AddUsedColumns(sheet.RootMap ?? ClassMap, sheet, usedColumns);

        var columns = new List<int>();
        var names = heading.ColumnNames;
        for (int i = 0; i < names.Count; i++)
        {
            var column = i + heading.ColumnOffset;
            if (!string.IsNullOrWhiteSpace(names[i]) && !usedColumns.Contains(column))
            {
                columns.Add(column);
            }
        }

        return [.. columns];
    }

    private static void AddUsedColumns(IMap map, SheetView sheet, HashSet<int> usedColumns)
        => MapTree.VisitCellMaps(map, null, (cellMap, _) =>
        {
            switch (cellMap)
            {
                case IValueMap valueMap:
                    AddUsedColumn(valueMap.Settings.Columns, sheet, usedColumns);
                    break;

                case ICellsMap { Settings.Columns: ExtraColumns }:
                    // Another extra columns member: it doesn't claim any columns.
                    break;

                case ICellsMap cellsMap:
                    try
                    {
                        if (cellsMap.Settings.Columns.FindColumns(sheet) is { } columns)
                        {
                            usedColumns.UnionWith(columns);
                        }
                    }
                    catch (ExcelMappingException)
                    {
                        // The member's columns can't be found, so it doesn't read any of the sheet's columns.
                    }

                    break;
            }
        });

    private static void AddUsedColumn(Columns columns, SheetView sheet, HashSet<int> usedColumns)
    {
        try
        {
            if (columns.FindColumn(sheet) is var column and not -1)
            {
                usedColumns.Add(column);
            }
        }
        catch (ExcelMappingException)
        {
            // The member's column can't be found (for example an optional column that is
            // missing), so it doesn't read any of the sheet's columns.
        }
    }
}

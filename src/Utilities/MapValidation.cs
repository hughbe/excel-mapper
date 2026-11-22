using ExcelMapper.Readers;

namespace ExcelMapper.Utilities;

/// <summary>
/// Finds the columns a map needs that a sheet doesn't have.
/// </summary>
internal static class MapValidation
{
    public static void FindMissingColumns(IMap map, ExcelSheet sheet, string? path, List<string> problems)
    {
        switch (map)
        {
            case ExcelClassMap classMap:
                if (classMap._valueMap is not null)
                {
                    FindMissingColumns(classMap._valueMap, sheet, path, problems);
                    return;
                }

                foreach (var property in classMap.Properties)
                {
                    FindMissingColumns(property.Map, sheet, Combine(path, property.Member.Name), problems);
                }

                return;

            case IValueMap valueMap:
                if (!valueMap.Optional)
                {
                    Check(valueMap.Columns, many: false, sheet, path, problems);
                }

                return;

            case IManyToOneMap cellsMap:
                if (!cellsMap.Optional)
                {
                    Check(cellsMap.Columns, many: true, sheet, path, problems);
                }

                return;

            case IEnumerableIndexerMap enumerableIndexerMap:
                foreach (var item in enumerableIndexerMap.Values)
                {
                    FindMissingColumns(item.Value, sheet, $"{path}[{item.Key}]", problems);
                }

                return;

            case IMultidimensionalIndexerMap multidimensionalIndexerMap:
                foreach (var item in multidimensionalIndexerMap.Values)
                {
                    FindMissingColumns(item.Value, sheet, $"{path}[{string.Join(", ", item.Key)}]", problems);
                }

                return;

            case IDictionaryIndexerMap dictionaryIndexerMap:
                foreach (var item in dictionaryIndexerMap.Values)
                {
                    FindMissingColumns(item.Value, sheet, $"{path}[\"{item.Key}\"]", problems);
                }

                return;
        }
    }

    /// <summary>
    /// Counts the columns that a map looks up by name, and how many of them the sheet's heading has.
    /// Used to find the heading row: columns found by index, and dictionaries of all columns, match
    /// any row.
    /// </summary>
    public static void CountNamedColumns(IMap map, ExcelSheet sheet, ref int lookedUp, ref int found)
    {
        switch (map)
        {
            case ExcelClassMap classMap:
                if (classMap._valueMap is not null)
                {
                    CountNamedColumns(classMap._valueMap, sheet, ref lookedUp, ref found);
                    return;
                }

                var properties = classMap.Properties;
                for (int i = 0; i < properties.Count; i++)
                {
                    CountNamedColumns(properties[i].Map, sheet, ref lookedUp, ref found);
                }

                return;

            case IValueMap valueMap when valueMap.Columns.LooksUpNames:
                lookedUp++;
                if (HasColumns(valueMap.Columns, many: false, sheet))
                {
                    found++;
                }

                return;

            case IManyToOneMap cellsMap when cellsMap.Columns.LooksUpNames:
                lookedUp++;
                if (HasColumns(cellsMap.Columns, many: true, sheet))
                {
                    found++;
                }

                return;

            case IEnumerableIndexerMap enumerableIndexerMap:
                foreach (var value in enumerableIndexerMap.Values.Values)
                {
                    CountNamedColumns(value, sheet, ref lookedUp, ref found);
                }

                return;

            case IMultidimensionalIndexerMap multidimensionalIndexerMap:
                foreach (var value in multidimensionalIndexerMap.Values.Values)
                {
                    CountNamedColumns(value, sheet, ref lookedUp, ref found);
                }

                return;

            case IDictionaryIndexerMap dictionaryIndexerMap:
                foreach (var value in dictionaryIndexerMap.Values.Values)
                {
                    CountNamedColumns(value, sheet, ref lookedUp, ref found);
                }

                return;
        }
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

    private static string Combine(string? path, string name) => path is null ? name : $"{path}.{name}";

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

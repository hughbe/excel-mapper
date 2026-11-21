using System.Linq;

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

            case IOneToOneMap oneToOneMap:
                if (!oneToOneMap.Optional)
                {
                    Check(() => oneToOneMap.ReaderFactory.GetCellReader(sheet) is not null, oneToOneMap.ReaderFactory, sheet, path, problems);
                }

                return;

            case IManyToOneMap manyToOneMap:
                if (!manyToOneMap.Optional)
                {
                    Check(() => manyToOneMap.ReaderFactory.GetCellsReader(sheet) is not null, manyToOneMap.ReaderFactory, sheet, path, problems);
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

    private static string Combine(string? path, string name) => path is null ? name : $"{path}.{name}";

    private static void Check(Func<bool> hasColumn, object readerFactory, ExcelSheet sheet, string? path, List<string> problems)
    {
        string? error = null;
        try
        {
            if (hasColumn())
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
        problems.Add(error is null ? $"{DescribeColumns(readerFactory, sheet)}{member}" : $"{error}{member}");
    }

    private static string DescribeColumns(object readerFactory, ExcelSheet sheet)
    {
        if (readerFactory is IColumnNameProviderCellReaderFactory nameProvider &&
            nameProvider.GetColumnName(sheet) is { Length: > 0 } columnName)
        {
            var hint = ExcelMappingException.GetSimilarColumnHint(sheet, columnName) is null
                ? string.Empty
                : $", did you mean \"{SimilarColumn(sheet, columnName)}\"?";
            return $"column \"{columnName}\"{hint}";
        }

        if (readerFactory is IColumnIndexProviderCellReaderFactory indexProvider &&
            indexProvider.GetColumnIndex(sheet) is int columnIndex and >= 0)
        {
            return $"column {ExcelRange.GetColumnLetters(columnIndex)} (index {columnIndex})";
        }

        if (readerFactory is IColumnNamesProviderCellReaderFactory namesProvider &&
            namesProvider.GetColumnNames(sheet) is { Count: > 0 } columnNames)
        {
            return $"one of columns {string.Join(", ", columnNames.Select(n => $"\"{n}\""))}";
        }

        if (readerFactory is IColumnIndicesProviderCellReaderFactory indicesProvider &&
            indicesProvider.GetColumnIndices(sheet) is { Count: > 0 } columnIndices)
        {
            return $"columns {string.Join(", ", columnIndices.Select(ExcelRange.GetColumnLetters))}";
        }

        return "a matching column";
    }

    private static string SimilarColumn(ExcelSheet sheet, string columnName)
    {
        var normalizedName = ColumnUtilities.RemoveSeparators(columnName);
        return sheet.Heading!.ColumnNames.First(c => string.Equals(ColumnUtilities.RemoveSeparators(c), normalizedName, StringComparison.OrdinalIgnoreCase));
    }
}

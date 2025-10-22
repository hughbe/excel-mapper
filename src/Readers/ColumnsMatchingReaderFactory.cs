namespace ExcelMapper.Readers;

/// <summary>
/// Reads the value of a multiples cells given a list of column names or a predicate matching the column name.
/// </summary>
public sealed class ColumnsMatchingReaderFactory : ICellReaderFactory, ICellsReaderFactory, IColumnIndicesProviderCellReaderFactory, IColumnNamesProviderCellReaderFactory
{
    public IExcelColumnMatcher Matcher { get; }

    /// <summary>
    /// Constructs a reader that reads the value of multiple cells given the predicate matching the column name.
    /// </summary>
    /// <param name="predicate">The predicate containing the column name to read.</param>
    public ColumnsMatchingReaderFactory(IExcelColumnMatcher matcher)
    {
        ArgumentNullException.ThrowIfNull(matcher);
        Matcher = matcher;
    }

    public ICellReader? GetCellReader(ExcelSheet sheet)
    {
        ArgumentNullException.ThrowIfNull(sheet);

        for (var columnIndex = 0; columnIndex < sheet.NumberOfColumns; columnIndex++)
        {
            if (Matcher.ColumnMatches(sheet, columnIndex))
            {
                return new ColumnIndexReader(columnIndex);
            }
        }

        return null;
    }

    public ICellsReader? GetCellsReader(ExcelSheet sheet)
    {
        ArgumentNullException.ThrowIfNull(sheet);

        var indices = new List<int>();
        for (var columnIndex = 0; columnIndex < sheet.NumberOfColumns; columnIndex++)
        {
            if (Matcher.ColumnMatches(sheet, columnIndex))
            {
                indices.Add(columnIndex);
            }
        }

        if (indices.Count == 0)
        {
            return null;
        }

        return new ColumnIndicesReader(indices);
    }

    public string[]? GetColumnNames(ExcelSheet sheet)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        if (sheet.Heading == null)
        {
            return null;
        }

        var names = new List<string>();
        for (var columnIndex = 0; columnIndex < sheet.NumberOfColumns; columnIndex++)
        {
            if (Matcher.ColumnMatches(sheet, columnIndex))
            {
                names.Add(sheet.Heading.GetColumnName(columnIndex)!);
            }
        }

        return names.ToArray();
    }

    public int[] GetColumnIndices(ExcelSheet sheet)
    {
        ArgumentNullException.ThrowIfNull(sheet);

        var indices = new List<int>();
        for (var columnIndex = 0; columnIndex < sheet.NumberOfColumns; columnIndex++)
        {
            if (Matcher.ColumnMatches(sheet, columnIndex))
            {
                indices.Add(columnIndex);
            }
        }

        return indices.ToArray();
    }
}
namespace ExcelMapper.Readers;

/// <summary>
/// Reads a multiple values of one or more columns given the name of each column.
/// </summary>
public sealed class ColumnNamesReaderFactory : ICellReaderFactory, ICellsReaderFactory, IColumnNamesProviderCellReaderFactory
{
    /// <summary>
    /// Gets the names of each column to read.
    /// </summary>
    public IReadOnlyList<string> ColumnNames { get; }

    /// <summary>
    /// Gets the string comparison used when matching column names.
    /// </summary>
    public StringComparison Comparison { get; } = StringComparison.OrdinalIgnoreCase;

    /// <summary>
    /// Constructs a reader that reads the values of one or more columns with a given name
    /// and returns the string value of for each column.
    /// </summary>
    /// <param name="columnNames">The names of each column to read.</param>
    public ColumnNamesReaderFactory(params IReadOnlyList<string> columnNames) : this(columnNames, StringComparison.OrdinalIgnoreCase)
    {
    }

    /// <summary>
    /// Constructs a reader that reads the values of one or more columns with a given name
    /// and returns the string value of for each column.
    /// </summary>
    /// <param name="columnNames">The names of each column to read.</param>
    public ColumnNamesReaderFactory(IReadOnlyList<string> columnNames, StringComparison comparison)
    {
        ColumnUtilities.ValidateColumnNames(columnNames);
        EnumUtilities.ValidateIsDefined(comparison);
        ColumnNames = columnNames;
        Comparison = comparison;
    }
    
    /// <inheritdoc/>
    public ICellReader? GetCellReader(ExcelSheet sheet)
    {
        foreach (var columnName in ColumnNames)
        {
            if (ColumnUtilities.TryGetColumnIndex(sheet, columnName, Comparison, out var columnIndex))
            {
                return new ColumnIndexReader(columnIndex);
            }
        }

        return null;
    }

    /// <inheritdoc/>
    public ICellsReader? GetCellsReader(ExcelSheet sheet)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        if (sheet.Heading == null)
        {
            throw new ExcelMappingException($"The sheet \"{sheet.Name}\" does not have a heading. Use a column index mapping instead.");
        }

        var indices = new int[ColumnNames.Count];
        for (int i = 0; i < ColumnNames.Count; i++)
        {
            if (!ColumnUtilities.TryGetColumnIndex(sheet, ColumnNames[i], Comparison, out var index))
            {
                return null;
            }

            indices[i] = index;
        }

        return new ColumnIndicesReader(indices);
    }

    /// <inheritdoc/>
    public IReadOnlyList<string> GetColumnNames(ExcelSheet sheet) => ColumnNames;
}

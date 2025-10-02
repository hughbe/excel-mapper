using System;
using ExcelMapper.Abstractions;
using ExcelMapper.Utilities;

namespace ExcelMapper.Readers;

/// <summary>
/// Reads the value of a single cell given a list of column names or a predicate matching the column name.
/// </summary>
public sealed class ColumnNameMatchingReaderFactory : ICellReaderFactory
{
    public string[]? ColumnNames { get; }
    public Func<string, bool>? Predicate { get; }

    /// <summary>
    /// Constructs a reader that reads the value of a single cell given the predicate matching the column name.
    /// </summary>
    /// <param name="predicate">The predicate containing the column name to read.</param>
    public ColumnNameMatchingReaderFactory(Func<string, bool> predicate)
    {
        Predicate = predicate ?? throw new ArgumentNullException(nameof(predicate));
    }
    /// <summary>
    /// Constructs a reader that reads the value of a single cell given a list of column names.
    /// </summary>
    /// <param name="columnNames">The list of column names to read.</param>
    public ColumnNameMatchingReaderFactory(params string[] columnNames)
    {
        ColumnNameUtilities.ValidateColumnNames(columnNames, nameof(columnNames));
        ColumnNames = columnNames;
    }

    public ICellReader? GetReader(ExcelSheet sheet)
    {
        if (sheet == null)
        {
            throw new ArgumentNullException(nameof(sheet));
        }
        if (sheet.Heading == null)
        {
            throw new ExcelMappingException($"The sheet \"{sheet.Name}\" does not have a heading. Use a column index mapping instead.");
        }

        if (ColumnNames != null)
        {
            foreach (string columnName in ColumnNames)
            {
                if (sheet.Heading.TryGetColumnIndex(columnName, out var columnIndex))
                {
                    return new ColumnIndexReader(columnIndex);
                }
            }
        }
        if (Predicate != null && sheet.Heading.TryGetFirstColumnMatchingIndex(Predicate, out var matchedColumnIndex))
        {
            return new ColumnIndexReader(matchedColumnIndex);
        }
        
        return null;
    }
}
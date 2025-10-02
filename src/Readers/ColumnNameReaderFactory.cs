using System;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Readers;

/// <summary>
/// Reads the value of a single cell given the name of it's column.
/// </summary>
public sealed class ColumnNameReaderFactory : ICellReaderFactory
{
    /// <summary>
    /// The name of the column to read.
    /// </summary>
    public string ColumnName { get; }

    /// <summary>
    /// Constructs a reader that reads the value of a single cell given the name of it's column.
    /// </summary>
    /// <param name="columnName">The name of the column to read.</param>
    public ColumnNameReaderFactory(string columnName)
    {
        if (columnName == null)
        {
            throw new ArgumentNullException(nameof(columnName));
        }

        if (columnName.Length == 0)
        {
            throw new ArgumentException("Column name cannot be empty.", nameof(columnName));
        }

        ColumnName = columnName;
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
        if (!sheet.Heading.TryGetColumnIndex(ColumnName, out int columnIndex))
        {
            return null;
        }

        return new ColumnIndexReader(columnIndex);
    }
}

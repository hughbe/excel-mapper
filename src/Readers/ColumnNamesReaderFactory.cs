using System;
using ExcelMapper.Abstractions;
using ExcelMapper.Utilities;

namespace ExcelMapper.Readers;

/// <summary>
/// Reads a multiple values of one or more columns given the name of each column.
/// </summary>
public sealed class ColumnNamesReaderFactory : ICellsReaderFactory
{
    /// <summary>
    /// Gets the names of each column to read.
    /// </summary>
    public string[] ColumnNames { get; }

    /// <summary>
    /// Constructs a reader that reads the values of one or more columns with a given name
    /// and returns the string value of for each column.
    /// </summary>
    /// <param name="columnNames">The names of each column to read.</param>
    public ColumnNamesReaderFactory(params string[] columnNames)
    {
        ColumnNameUtilities.ValidateColumnNames(columnNames, nameof(columnNames));
        ColumnNames = columnNames;
    }

    public ICellsReader? GetReader(ExcelSheet sheet)
    {
        if (sheet == null)
        {
            throw new ArgumentNullException(nameof(sheet));
        }
        if (sheet.Heading == null)
        {
            throw new ExcelMappingException($"The sheet \"{sheet.Name}\" does not have a heading. Use a column index mapping instead.");
        }

        var indices = new int[ColumnNames.Length];
        for (int i = 0; i < ColumnNames.Length; i++)
        {
            if (!sheet.Heading.TryGetColumnIndex(ColumnNames[i], out int index))
            {
                return null;
            }

            indices[i] = index;
        }

        return new ColumnIndicesReader(indices);
    }
}

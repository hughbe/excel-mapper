using System.Diagnostics.CodeAnalysis;
using ExcelDataReader;

namespace ExcelMapper.Readers;

/// <summary>
/// Reads a multiple values of all columns in a sheet.
/// </summary>
public sealed class AllColumnNamesReaderFactory : ICellsReaderFactory
{
    /// <inheritdoc/>
    public ICellsReader? GetCellsReader(ExcelSheet sheet)
    {
        ThrowHelpers.ThrowIfNull(sheet, nameof(sheet));
        if (sheet.Heading == null)
        {
            throw new ExcelMappingException($"The sheet \"{sheet.Name}\" does not have a heading. Use a column index map instead.");
        }

        var columnNames = sheet.Heading.ColumnNames;
        var count = 0;

        // First pass: count non-empty column names
        for (int i = 0; i < columnNames.Count; i++)
        {
            if (!string.IsNullOrWhiteSpace(columnNames[i]))
            {
                count++;
            }
        }

        // Second pass: collect indices
        var indices = new int[count];
        int index = 0;
        for (int i = 0; i < columnNames.Count; i++)
        {
            var columnName = columnNames[i];
            if (!string.IsNullOrWhiteSpace(columnName))
            {
                indices[index++] = sheet.Heading.GetColumnIndex(columnName);
            }
        }

        return new ColumnIndicesReader(indices);
    }
}

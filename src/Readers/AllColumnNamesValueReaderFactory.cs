using System;
using System.Linq;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Readers;

/// <summary>
/// Reads a multiple values of all columns in a sheet.
/// </summary>
public sealed class AllColumnNamesReaderFactory : ICellsReaderFactory
{
    public ICellsReader? GetCellsReader(ExcelSheet sheet)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        if (sheet.Heading == null)
        {
            throw new ExcelMappingException($"The sheet \"{sheet.Name}\" does not have a heading. Use a column index map instead.");
        }

        var indices = sheet.Heading.ColumnNames
            .Where(s => !string.IsNullOrWhiteSpace(s))
            .Select(columnName => sheet.Heading.GetColumnIndex(columnName))
            .ToArray();
        return new ColumnIndicesReader(indices);
    }
}

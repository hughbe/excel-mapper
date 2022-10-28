using ExcelDataReader;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Readers;
 
/// <summary>
/// Reads a multiple values of all columns in a sheet.
/// </summary>
public sealed class AllColumnNamesValueReader : IMultipleCellReader
{
    public bool TryGetCells(ExcelRow row, IExcelDataReader reader, out IEnumerable<ExcelCell> cells)
    {
        if (row.Sheet.Heading == null)
        {
            throw new ExcelMappingException($"The sheet \"{row.Sheet.Name}\" does not have a heading. Use a column index mapping instead.");
        }

        cells = row.Sheet.Heading.ColumnNames
            .Where(s => !string.IsNullOrWhiteSpace(s))
            .Select(columnName =>
            {
                var columnIndex = row.Sheet.Heading.GetColumnIndex(columnName);
                return new ExcelCell(row.Sheet, row.RowIndex, columnIndex);
            });
        return true;
    }
}

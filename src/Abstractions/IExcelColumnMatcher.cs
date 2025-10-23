namespace ExcelMapper.Abstractions;

/// <summary>
/// Matches Excel columns based on custom criteria.
/// </summary>
public interface IExcelColumnMatcher
{
    /// <summary>
    /// Determines if the specified column matches the criteria.
    /// </summary>
    /// <param name="sheet">The Excel sheet.</param>
    /// <param name="columnIndex">The index of the column to check.</param>
    /// <returns>True if the column matches the criteria; otherwise, false.</returns>
    bool ColumnMatches(ExcelSheet sheet, int columnIndex);
}

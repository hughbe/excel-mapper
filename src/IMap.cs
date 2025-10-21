using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using ExcelDataReader;

namespace ExcelMapper;

/// <summary>
/// Represents a mapping strategy that reads one or more cells from an Excel sheet and converts them to an object.
/// </summary>
public interface IMap
{
    /// <summary>
    /// Attempts to read and map cell value(s) from the current row to an object.
    /// </summary>
    /// <param name="sheet">The Excel sheet being read, containing metadata like column names and current position.</param>
    /// <param name="rowIndex">The zero-based index of the row being mapped (relative to the beginning of the sheet, including header).</param>
    /// <param name="reader">The underlying Excel data reader providing access to cell values in the current row.</param>
    /// <param name="member">The property or field being mapped to, if applicable. Used for error messages and attribute inspection. Can be null for top-level mappings.</param>
    /// <param name="value">When this method returns true, contains the mapped object. When false, contains null.</param>
    /// <returns>True if the mapping succeeded and a value was produced; false if the mapping failed or no value could be produced (e.g., optional column not found).</returns>
    bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MemberInfo? member, [NotNullWhen(true)] out object? value);
}

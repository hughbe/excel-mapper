using System.Diagnostics.CodeAnalysis;
using ExcelDataReader;

namespace ExcelMapper.Abstractions;

/// <summary>
/// Tries to read the values of multiple cells.
/// </summary>
public interface ICellsReader
{
    /// <summary>
    /// Tries to read the value of multiple cells.
    /// </summary>
    bool TryGetValues(IExcelDataReader reader, bool preserveFormatting, [NotNullWhen(true)] out IEnumerable<ReadCellResult>? result);
}

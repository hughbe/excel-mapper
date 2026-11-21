using ExcelDataReader;

namespace ExcelMapper.Abstractions;

/// <summary>
/// Tries to read the value of a single cell.
/// </summary>
internal interface ICellReader
{
    
    /// <summary>
    /// Tries to read the value of a single cell.
    /// </summary>
    bool TryGetValue(IExcelDataReader reader, bool preserveFormatting, out ReadCellResult result);
}

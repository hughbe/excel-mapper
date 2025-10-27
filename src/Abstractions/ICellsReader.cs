using System.Diagnostics.CodeAnalysis;
using ExcelDataReader;

namespace ExcelMapper.Abstractions;

/// <summary>
/// Tries to read the values of multiple cells.
/// </summary>
public interface ICellsReader
{
    bool Start(IExcelDataReader reader, bool preserveFormatting, out int count);

    bool TryGetNext([NotNullWhen(true)] out ReadCellResult result);

    void Reset();
}

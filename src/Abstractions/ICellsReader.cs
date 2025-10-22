using System.Diagnostics.CodeAnalysis;
using ExcelDataReader;

namespace ExcelMapper.Abstractions;

/// <summary>
/// An interface that describes an object that can read the values of multiple cells in a row.
/// This describes a one-to-many or many-to-many mapping between a cell and a mapped value.
/// </summary>
public interface ICellsReader
{
    bool TryGetValues(IExcelDataReader reader, bool preserveFormatting, [NotNullWhen(true)] out IEnumerable<ReadCellResult>? result);
}

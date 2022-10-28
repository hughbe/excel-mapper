using ExcelDataReader;

namespace ExcelMapper.Abstractions;
 
/// <summary>
/// An interface that describes an object that can read the values of multiple cells in a row.
/// This describes a many-to-many mapping between a property and values.
/// </summary>
public interface IMultipleCellReader
{
    bool TryGetCells(ExcelRow row, IExcelDataReader reader, out IEnumerable<ExcelCell> cells);
}

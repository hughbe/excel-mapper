using ExcelDataReader;

namespace ExcelMapper.Abstractions
{
    /// <summary>
    /// An interface that describes an object that can read the value of a single cell in a row.
    /// This describes a one-to-one mapping between a cell and a mapped value.
    /// </summary>
    public interface ISingleCellValueReader
    {
        bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, out ReadCellValueResult result);
    }
}

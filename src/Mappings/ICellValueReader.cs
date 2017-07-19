using ExcelDataReader;

namespace ExcelMapper.Mappings
{
    /// <summary>
    /// An interface that describes an object that can read the value of a single cell in a row.
    /// This describes a 1-1 mapping between a cell and a mapped value.
    /// </summary>
    public interface ICellValueReader
    {
        ReadCellValueResult GetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader);
    }
}

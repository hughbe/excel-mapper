using ExcelDataReader;

namespace ExcelMapper.Mappings
{
    public interface ISinglePropertyMapper
    {
        int GetColumnIndex(ExcelSheet sheet, int rowIndex, IExcelDataReader reader);
    }
}

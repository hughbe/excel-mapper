using ExcelDataReader;

namespace ExcelMapper.Mappings
{
    public interface ISinglePropertyMapper
    {
        MapResult GetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader);
    }
}

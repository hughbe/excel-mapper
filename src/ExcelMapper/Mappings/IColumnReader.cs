using ExcelDataReader;

namespace ExcelMapper.Mappings
{
    public interface ISingleValueReader
    {
        ReadResult GetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader);
    }
}

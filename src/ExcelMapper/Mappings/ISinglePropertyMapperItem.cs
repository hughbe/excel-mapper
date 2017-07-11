using ExcelDataReader;

namespace ExcelMapper.Mappings
{
    public interface ISinglePropertyMappingItem
    {
        PropertyMappingResult GetProperty(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, int columnIndex, string stringValue);
    }
}

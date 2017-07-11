using ExcelDataReader;

namespace ExcelMapper.Mappings.Items
{
    internal class ParseAsStringMappingItem : ISinglePropertyMappingItem
    {
        public PropertyMappingResult GetProperty(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, int columnIndex, string stringValue)
        {
            return PropertyMappingResult.Success(stringValue);
        }
    }
}

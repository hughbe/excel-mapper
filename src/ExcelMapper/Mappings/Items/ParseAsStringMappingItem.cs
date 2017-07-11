using ExcelDataReader;

namespace ExcelMapper.Mappings.Items
{
    internal class ParseAsStringMappingItem : ISinglePropertyMappingItem
    {
        public PropertyMappingResult GetProperty(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MapResult result)
        {
            return PropertyMappingResult.Began(result.StringValue);
        }
    }
}

using ExcelDataReader;

namespace ExcelMapper.Mappings.Items
{
    internal class ParseAsBoolMappingItem : ISinglePropertyMappingItem
    {
        public PropertyMappingResult GetProperty(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, int columnIndex, string stringValue)
        {
            // Excel transforms bool values such as "true" or "false" to "1" or "0".
            if (stringValue == "1")
            {
                return PropertyMappingResult.Success(true);
            }
            else if (stringValue == "0")
            {
                return PropertyMappingResult.Success(false);
            }

            if (!bool.TryParse(stringValue, out bool result))
            {
                return PropertyMappingResult.Invalid();
            }

            return PropertyMappingResult.Success(result);
        }
    }
}

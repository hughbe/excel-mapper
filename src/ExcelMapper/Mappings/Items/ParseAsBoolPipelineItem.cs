using ExcelDataReader;

namespace ExcelMapper.Mappings.Items
{
    public class ParseAsBoolMappingItem : ISinglePropertyMappingItem
    {
        public PropertyMappingResult GetProperty(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, ReadResult mapResult)
        {
            // Excel transforms bool values such as "true" or "false" to "1" or "0".
            if (mapResult.StringValue == "1")
            {
                return PropertyMappingResult.Success(true);
            }
            else if (mapResult.StringValue == "0")
            {
                return PropertyMappingResult.Success(false);
            }

            if (!bool.TryParse(mapResult.StringValue, out bool result))
            {
                return PropertyMappingResult.Invalid();
            }

            return PropertyMappingResult.Success(result);
        }
    }
}

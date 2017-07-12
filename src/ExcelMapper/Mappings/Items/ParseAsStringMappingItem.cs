using ExcelDataReader;

namespace ExcelMapper.Mappings.Items
{
    public class ParseAsStringMappingItem : ISinglePropertyMappingItem
    {
        public PropertyMappingResult GetProperty(ReadResult result)
        {
            return PropertyMappingResult.Began(result.StringValue);
        }
    }
}

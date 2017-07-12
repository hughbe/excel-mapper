namespace ExcelMapper.Mappings.Mappers
{
    public class BoolMapper : IStringValueMapper
    {
        public PropertyMappingResultType GetProperty(ReadResult readResult, ref object value)
        {
            // Excel transforms bool values such as "true" or "false" to "1" or "0".
            if (readResult.StringValue == "1")
            {
                value = true;
                return PropertyMappingResultType.Success;
            }
            else if (readResult.StringValue == "0")
            {
                value = false;
                return PropertyMappingResultType.Success;
            }

            if (!bool.TryParse(readResult.StringValue, out bool result))
            {
                return PropertyMappingResultType.Invalid;
            }

            value = result;
            return PropertyMappingResultType.Success;
        }
    }
}

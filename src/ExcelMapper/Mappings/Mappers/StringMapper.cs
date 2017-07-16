namespace ExcelMapper.Mappings.Mappers
{
    public class StringMapper : IStringValueMapper
    {
        public PropertyMappingResultType GetProperty(ReadResult result, ref object value)
        {
            value = result.StringValue;
            return PropertyMappingResultType.SuccessIfNoOtherSuccess;
        }
    }
}

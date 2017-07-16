using System;

namespace ExcelMapper.Mappings.Mappers
{
    public class UriMapper : IStringValueMapper
    {
        public PropertyMappingResultType GetProperty(ReadResult readResult, ref object value)
        {
            if (!Uri.TryCreate(readResult.StringValue, UriKind.Absolute, out Uri result))
            {
                return PropertyMappingResultType.Invalid;
            }

            value = result;
            return PropertyMappingResultType.Success;
        }
    }
}

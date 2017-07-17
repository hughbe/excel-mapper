using System;

namespace ExcelMapper.Mappings.Mappers
{
    /// <summary>
    /// Tries to map the value of a cell to an absolute Uri.
    /// </summary>
    public class UriMapper : ICellValueMapper
    {
        public PropertyMappingResultType GetProperty(ReadCellValueResult readResult, ref object value)
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

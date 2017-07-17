using System;

namespace ExcelMapper.Mappings.Mappers
{
    /// <summary>
    /// Tries to map the value of a cell to an absolute Uri.
    /// </summary>
    public class UriMapper : ICellValueMapper
    {
        public PropertyMapperResultType GetProperty(ReadCellValueResult readResult, ref object value)
        {
            if (!Uri.TryCreate(readResult.StringValue, UriKind.Absolute, out Uri result))
            {
                return PropertyMapperResultType.Invalid;
            }

            value = result;
            return PropertyMapperResultType.Success;
        }
    }
}

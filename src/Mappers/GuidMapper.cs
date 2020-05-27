using System;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Mappers
{
    /// <summary>
    /// A mapper that tries to map the value of a cell to a guid.
    /// </summary>
    public class GuidMapper : ICellValueMapper
    {
        public PropertyMapperResultType MapCellValue(ReadCellValueResult readResult, ref object value)
        {
            if (!Guid.TryParse(readResult.StringValue, out Guid result))
            {
                return PropertyMapperResultType.Invalid;
            }

            value = result;
            return PropertyMapperResultType.Success;
        }
    }
}

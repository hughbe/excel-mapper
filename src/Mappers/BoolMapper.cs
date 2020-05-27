using ExcelMapper.Abstractions;

namespace ExcelMapper.Mappers
{
    /// <summary>
    /// A mapper that tries to map the value of a cell to a bool.
    /// </summary>
    public class BoolMapper : ICellValueMapper
    {
        public PropertyMapperResultType MapCellValue(ReadCellValueResult readResult, ref object value)
        {
            // Excel transforms bool values such as "true" or "false" to "1" or "0".
            if (readResult.StringValue == "1")
            {
                value = true;
                return PropertyMapperResultType.Success;
            }

            if (readResult.StringValue == "0")
            {
                value = false;
                return PropertyMapperResultType.Success;
            }

            if (!bool.TryParse(readResult.StringValue, out bool result))
            {
                return PropertyMapperResultType.Invalid;
            }

            value = result;
            return PropertyMapperResultType.Success;
        }
    }
}

using System;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Mappers
{
    /// <summary>
    /// A mapper that tries to map the value of a cell to a bool.
    /// </summary>
    public class BoolMapper : ICellValueMapper
    {
        private static object s_boxedTrue = true;
        private static object s_boxedFalse = false;

        public CellValueMapperResult MapCellValue(ReadCellValueResult readResult)
        {
            // Excel transforms bool values such as "true" or "false" to "1" or "0".
            if (readResult.StringValue == "1")
            {
                return CellValueMapperResult.Success(s_boxedTrue);
            }
            else if (readResult.StringValue == "0")
            {
                return CellValueMapperResult.Success(s_boxedFalse);
            }

            try
            {
                bool result = bool.Parse(readResult.StringValue);
                return CellValueMapperResult.Success(result);
            }
            catch (Exception exception)
            {
                return CellValueMapperResult.Invalid(exception);
            }
        }
    }
}

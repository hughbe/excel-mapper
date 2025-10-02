using System;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Mappers
{
    /// <summary>
    /// A mapper that tries to map the value of a cell to a bool.
    /// </summary>
    public class BoolMapper : ICellMapper
    {
        private static readonly object s_boxedTrue = true;
        private static readonly object s_boxedFalse = false;

        public CellMapperResult MapCellValue(ReadCellResult readResult)
        {
            // Excel transforms bool values such as "true" or "false" to "1" or "0".
            if (readResult.StringValue == "1")
            {
                return CellMapperResult.Success(s_boxedTrue);
            }
            if (readResult.StringValue == "0")
            {
                return CellMapperResult.Success(s_boxedFalse);
            }

            try
            {
                // Discarding readResult.StringValue nullability warning.
                // If null - CellMapperResult.Invalid with ArgumentNullException will be returned
                bool result = bool.Parse(readResult.StringValue!);
                return CellMapperResult.Success(result);
            }
            catch (Exception exception)
            {
                return CellMapperResult.Invalid(exception);
            }
        }
    }
}

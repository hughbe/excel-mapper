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
            if (readResult.Reader != null && readResult.Reader.GetValue(readResult.ColumnIndex) is bool boolValue)
            {
                return CellMapperResult.Success(boolValue);
            }
            
            var stringValue = readResult.GetString();

            // Excel transforms bool values such as "true" or "false" to "1" or "0".
            if (stringValue == "1")
            {
                return CellMapperResult.Success(s_boxedTrue);
            }
            if (stringValue == "0")
            {
                return CellMapperResult.Success(s_boxedFalse);
            }

            try
            {
                // Discarding stringValue nullability warning.
                // If null - CellMapperResult.Invalid with ArgumentNullException will be returned
                var result = bool.Parse(stringValue!);
                return CellMapperResult.Success(result);
            }
            catch (Exception exception)
            {
                return CellMapperResult.Invalid(exception);
            }
        }
    }
}

using System;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Mappers
{
    /// <summary>
    /// A mapper that tries to map the value of a cell to a guid.
    /// </summary>
    public class GuidMapper : ICellValueMapper
    {
        public CellValueMapperResult MapCellValue(ReadCellValueResult readResult)
        {
            try
            {
                // Discarding readResult.StringValue nullability warning.
                // If null - CellValueMapperResult.Invalid with ArgumentNullException will be returned
                Guid result = Guid.Parse(readResult.StringValue!);
                return CellValueMapperResult.Success(result);
            }
            catch (Exception exception)
            {
                return CellValueMapperResult.Invalid(exception);
            }
        }
    }
}

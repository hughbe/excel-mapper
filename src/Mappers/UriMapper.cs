using System;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Mappers
{
    /// <summary>
    /// Tries to map the value of a cell to an absolute Uri.
    /// </summary>
    public class UriMapper : ICellValueMapper
    {
        public CellValueMapperResult MapCellValue(ReadCellValueResult readResult)
        {
            try
            {
                // Discarding readResult.StringValue nullability warning.
                // If null - CellValueMapperResult.Invalid with ArgumentNullException will be returned
                var uri = new Uri(readResult.StringValue!, UriKind.Absolute);
                return CellValueMapperResult.Success(uri);
            }
            catch (Exception exception)
            {
                return CellValueMapperResult.Invalid(exception);
            }
        }
    }
}

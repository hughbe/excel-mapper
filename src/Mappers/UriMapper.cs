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
                var uri = new Uri(readResult.StringValue, UriKind.Absolute);
                return CellValueMapperResult.Success(uri);
            }
            catch (Exception exception)
            {
                return CellValueMapperResult.Invalid(exception);
            }
        }
    }
}

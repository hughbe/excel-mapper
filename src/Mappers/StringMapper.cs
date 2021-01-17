using ExcelMapper.Abstractions;

namespace ExcelMapper.Mappers
{
    /// <summary>
    /// A mapper that returns the string value of a cell.
    /// </summary>
    public class StringMapper : ICellValueMapper
    {
        public CellValueMapperResult MapCellValue(ReadCellValueResult result)
        {
            return CellValueMapperResult.SuccessIfNoOtherSuccess(result.StringValue);
        }
    }
}

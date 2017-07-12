using ExcelDataReader;

namespace ExcelMapper.Mappings.Fallbacks
{
    public class ThrowFallback : ISinglePropertyMappingItem
    {
        public PropertyMappingResult GetProperty(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, ReadResult mapResult)
        {
            throw new ExcelMappingException($"Invalid result for parameter \"{mapResult.StringValue}\" of type \"TODO\"", sheet, rowIndex, mapResult.ColumnIndex);
        }
    }
}

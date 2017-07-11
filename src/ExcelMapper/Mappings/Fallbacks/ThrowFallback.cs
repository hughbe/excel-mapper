using ExcelDataReader;

namespace ExcelMapper.Mappings.Fallbacks
{
    internal class ThrowFallback : ISinglePropertyMappingItem
    {
        public PropertyMappingResult GetProperty(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MapResult mapResult)
        {
            throw new ExcelMappingException($"Invalid result for parameter \"{mapResult.StringValue}\" of type \"TODO\"", sheet, rowIndex, mapResult.ColumnIndex);
        }
    }
}

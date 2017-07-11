using ExcelDataReader;

namespace ExcelMapper.Mappings.Fallbacks
{
    internal class ThrowFallback : ISinglePropertyMappingItem
    {
        public PropertyMappingResult GetProperty(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, int columnIndex, string stringValue)
        {
            throw new ExcelMappingException($"Invalid result for parameter \"{stringValue}\" of type \"TODO\"", sheet, rowIndex, columnIndex);
        }
    }
}

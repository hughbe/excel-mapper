namespace ExcelMapper.Mappings.Fallbacks
{
    public class ThrowFallback : IFallbackItem
    {
        public object PerformFallback(ExcelSheet sheet, int rowIndex, ReadResult mapResult)
        {
            throw new ExcelMappingException($"Invalid result for parameter \"{mapResult.StringValue}\" of type \"TODO\"", sheet, rowIndex, mapResult.ColumnIndex);
        }
    }
}

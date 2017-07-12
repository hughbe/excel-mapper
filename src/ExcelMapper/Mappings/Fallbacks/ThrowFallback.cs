namespace ExcelMapper.Mappings.Fallbacks
{
    public class ThrowFallback : IFallbackItem
    {
        public object PerformFallback(ExcelSheet sheet, int rowIndex, ReadResult readResult)
        {
            throw new ExcelMappingException($"Invalid result for parameter \"{readResult.StringValue}\" of type \"TODO\"", sheet, rowIndex, readResult.ColumnIndex);
        }
    }
}

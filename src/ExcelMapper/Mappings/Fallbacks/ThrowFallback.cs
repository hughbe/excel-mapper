namespace ExcelMapper.Mappings.Fallbacks
{
    /// <summary>
    /// A fallback that throws an ExcelMappingException when falling back.
    /// </summary>
    public class ThrowFallback : IFallbackItem
    {
        public object PerformFallback(ExcelSheet sheet, int rowIndex, ReadCellValueResult readResult)
        {
            throw new ExcelMappingException($"Invalid result for parameter \"{readResult.StringValue}\" of type \"TODO\"", sheet, rowIndex, readResult.ColumnIndex);
        }
    }
}

namespace ExcelMapper.Mappings
{
    public interface IFallbackItem
    {
        object PerformFallback(ExcelSheet sheet, int rowIndex, ReadResult result);
    }
}

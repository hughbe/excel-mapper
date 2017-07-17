namespace ExcelMapper.Mappings
{
    /// <summary>
    /// An interface that describes an object that is called when an empty or invalid
    /// value of cell is encountered. This can be used to return a fixed value or to throw
    /// an exception.
    /// TODO: pass information about the member.
    /// </summary>
    public interface IFallbackItem
    {
        object PerformFallback(ExcelSheet sheet, int rowIndex, ReadCellValueResult readResult);
    }
}

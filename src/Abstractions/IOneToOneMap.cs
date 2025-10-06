namespace ExcelMapper.Abstractions;

public interface IOneToOneMap : IToOneMap
{
    /// <summary>
    /// Gets or sets the factory that creates a reader for a cell value.
    /// </summary>
    public ICellReaderFactory ReaderFactory { get; set; }
}

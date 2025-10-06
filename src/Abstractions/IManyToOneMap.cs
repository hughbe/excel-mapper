namespace ExcelMapper.Abstractions;

public interface IManyToOneMap : IToOneMap
{
    /// <summary>
    /// Gets or sets the factory that creates a reader for cell values.
    /// </summary>
    public ICellsReaderFactory ReaderFactory { get; set; }
}

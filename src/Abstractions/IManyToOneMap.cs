namespace ExcelMapper.Abstractions;

/// <summary>
/// Map for many-to-one relationships, for example, mapping multiple cells to a single object.
/// </summary>
public interface IManyToOneMap : IToOneMap
{
    /// <summary>
    /// Gets or sets the factory that creates a reader for cell values.
    /// </summary>
    public ICellsReaderFactory ReaderFactory { get; set; }
}

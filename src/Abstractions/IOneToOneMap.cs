namespace ExcelMapper.Abstractions;

/// <summary>
/// Map for one-to-one relationships, for example, mapping a single cell to a single object.
/// </summary>
public interface IOneToOneMap : IToOneMap
{
    /// <summary>
    /// Gets or sets the factory that creates a reader for a cell value.
    /// </summary>
    public ICellReaderFactory ReaderFactory { get; set; }
}

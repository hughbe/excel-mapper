namespace ExcelMapper.Abstractions;

/// <summary>
/// Map for one-to-one relationships, for example, mapping a single cell to a single object.
/// </summary>
internal interface IValueMap : IToOneMap
{
    /// <summary>
    /// Gets or sets the factory that creates a reader for a cell value.
    /// </summary>
    /// <summary>
    /// The columns the map reads.
    /// </summary>
    public Readers.Columns Columns { get; }
}

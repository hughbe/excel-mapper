namespace ExcelMapper.Abstractions;

/// <summary>
/// Represents a map from a single Excel cell to a value.
/// </summary>
public interface IOneToOneMap<out T> : IMap<T>
{
    /// <summary>
    /// Gets or sets the <see cref="ICellReader"/> used to read the cell value.
    /// </summary>
    ICellReader Reader { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether the cell value is optional.
    /// </summary>
    bool Optional { get; set; }

    /// <summary>
    /// Gets the list of <see cref="ICellMapper"/>s used to convert the cell value.
    /// </summary>

    CellValueMapperCollection Mappers { get; }
}

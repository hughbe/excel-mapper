namespace ExcelMapper.Abstractions;

/// <summary>
/// A map that reads cells from columns: a single value (<see cref="IValueMap"/>), or a collection
/// or dictionary (<see cref="ICellsMap"/>).
/// </summary>
internal interface IColumnsMap : IMap
{
    /// <summary>
    /// The columns the map reads.
    /// </summary>
    Readers.Columns Columns { get; set; }

    /// <summary>
    /// Whether a missing column leaves the member's value, rather than being an error.
    /// </summary>
    bool Optional { get; set; }

    /// <summary>
    /// Whether cells are read as the text Excel shows, with their number format applied.
    /// </summary>
    bool PreserveFormatting { get; set; }

    /// <summary>
    /// How each cell is converted: the map itself for a single value, or its elements' or values'
    /// <see cref="ElementMap{T}"/>.
    /// </summary>
    ValuePipeline Pipeline { get; }
}

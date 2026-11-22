namespace ExcelMapper.Abstractions;

/// <summary>
/// A map that reads cells from columns: a single value (<see cref="ValueMap{T}"/>), or a collection
/// or dictionary (<see cref="ICellsMap"/>).
/// </summary>
internal interface IColumnsMap : IMap
{
    /// <summary>
    /// Where the map reads its cells, and how: its columns, whether they can be missing and
    /// whether cells are read as formatted text.
    /// </summary>
    Readers.ColumnSettings Settings { get; }

    /// <summary>
    /// How each cell is converted: the map itself for a single value, or its elements' or values'
    /// <see cref="ElementMap{T}"/>.
    /// </summary>
    ValuePipeline Pipeline { get; }
}

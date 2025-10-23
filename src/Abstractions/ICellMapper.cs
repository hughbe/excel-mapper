namespace ExcelMapper.Abstractions;

/// <summary>
/// Tries to map the value of a cell to an object.
/// </summary>
public interface ICellMapper
{
    /// <summary>
    /// Maps the value of a cell to an object.
    /// </summary>
    CellMapperResult Map(ReadCellResult readResult);
}

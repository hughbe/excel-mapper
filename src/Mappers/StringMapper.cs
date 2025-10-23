namespace ExcelMapper.Mappers;

/// <summary>
/// A mapper that returns the string value of a cell.
/// </summary>
public class StringMapper : ICellMapper
{
    /// <inheritdoc/>
    public CellMapperResult Map(ReadCellResult result)
        => CellMapperResult.SuccessIfNoOtherSuccess(result.GetString());
}

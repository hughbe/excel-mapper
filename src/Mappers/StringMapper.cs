namespace ExcelMapper.Mappers;

/// <summary>
/// A mapper that returns the string value of a cell.
/// </summary>
public class StringMapper : ICellMapper, ICellConverter<string?>
{
    /// <inheritdoc/>
    public CellMapperResult Map(ReadCellResult result)
        => CellMapperResult.SuccessIfNoOtherSuccess(result.GetString());

    CellConversion ICellConverter<string?>.Convert(ReadCellResult cell, out string? value)
    {
        value = cell.GetString();
        return CellConversion.SuccessUnlessReplaced;
    }
}

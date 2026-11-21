namespace ExcelMapper.Mappers;

/// <summary>
/// A mapper that returns the string value of a cell.
/// </summary>
public class StringMapper : ICellMapper, ITypedCellMapper<string?>
{
    /// <inheritdoc/>
    public CellMapperResult Map(ReadCellResult result)
        => CellMapperResult.SuccessIfNoOtherSuccess(result.GetString());

    CellConversion ITypedCellMapper<string?>.Convert(ReadCellResult cell, out string? value)
    {
        value = cell.GetString();
        return CellConversion.SuccessUnlessReplaced;
    }
}

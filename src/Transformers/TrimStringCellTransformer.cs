namespace ExcelMapper.Transformers;

/// <summary>
/// Trims whitespace from string cell values.
/// </summary>
public class TrimStringCellTransformer : Abstractions.ICellTransformer
{
    /// <inheritdoc/>
    public string? TransformStringValue(ExcelSheet sheet, int rowIndex, ReadCellResult readResult)
        => readResult.GetString()?.Trim();
}

using ExcelMapper.Abstractions;

namespace ExcelMapper.Transformers;

public class TrimCellTransformer : ICellTransformer
{
    public string? TransformStringValue(ExcelSheet sheet, int rowIndex, ReadCellResult readResult)
        => readResult.GetString()?.Trim();
}

using ExcelMapper.Abstractions;

namespace ExcelMapper.Transformers;

public class TrimCellValueTransformer : ICellTransformer
{
    public string? TransformStringValue(ExcelSheet sheet, int rowIndex, ReadCellResult readResult)
        => readResult.StringValue?.Trim();
}

namespace ExcelMapper.Abstractions;

public interface ICellTransformer
{
    string? TransformStringValue(ExcelSheet sheet, int rowIndex, ReadCellResult stringValue);
}

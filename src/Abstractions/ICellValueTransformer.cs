namespace ExcelMapper.Abstractions
{
    public interface ICellValueTransformer
    {
        string TransformStringValue(ExcelSheet sheet, int rowIndex, ReadCellValueResult stringValue);
    }
}

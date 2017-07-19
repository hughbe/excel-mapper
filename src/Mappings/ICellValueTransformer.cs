namespace ExcelMapper.Mappings
{
    public interface ICellValueTransformer
    {
        string TransformStringValue(ExcelSheet sheet, int rowIndex, ReadCellValueResult stringValue);
    }
}

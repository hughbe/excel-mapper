namespace ExcelMapper.Mappings
{
    public interface IStringValueTransformer
    {
        string TransformStringValue(ExcelSheet sheet, int rowIndex, ReadResult stringValue);
    }
}

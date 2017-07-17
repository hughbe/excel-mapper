namespace ExcelMapper.Mappings.Transformers
{
    public class TrimCellValueTransformer : ICellValueTransformer
    {
        public string TransformStringValue(ExcelSheet sheet, int rowIndex, ReadCellValueResult readResult)
        {
            if (readResult.StringValue == null)
            {
                return readResult.StringValue;
            }

            return readResult.StringValue.Trim();
        }
    }
}

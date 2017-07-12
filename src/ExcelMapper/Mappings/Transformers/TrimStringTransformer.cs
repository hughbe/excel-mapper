namespace ExcelMapper.Mappings.Transformers
{
    public class TrimStringTransformer : IStringValueTransformer
    {
        public string TransformStringValue(ExcelSheet sheet, int rowIndex, ReadResult readResult)
        {
            if (readResult.StringValue == null)
            {
                return readResult.StringValue;
            }

            return readResult.StringValue.Trim();
        }
    }
}

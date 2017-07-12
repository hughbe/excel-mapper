using ExcelDataReader;

namespace ExcelMapper.Mappings.Transformers
{
    public class TrimStringTransformer : IStringValueTransformer
    {
        public string TransformStringValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, ReadResult mapResult)
        {
            if (mapResult.StringValue == null)
            {
                return mapResult.StringValue;
            }

            return mapResult.StringValue.Trim();
        }
    }
}

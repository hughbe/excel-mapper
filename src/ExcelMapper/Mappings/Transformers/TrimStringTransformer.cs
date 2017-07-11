using ExcelDataReader;

namespace ExcelMapper.Mappings.Transformers
{
    internal class TrimStringTransformer : IStringValueTransformer
    {
        public string TransformStringValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, int columnIndex, string stringValue)
        {
            if (stringValue == null)
            {
                return stringValue;
            }

            return stringValue.Trim();
        }
    }
}

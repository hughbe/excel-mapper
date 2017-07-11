using ExcelDataReader;

namespace ExcelMapper.Mappings.Fallbacks
{
    internal class FixedValueFallback : ISinglePropertyMappingItem
    {
        public object Value { get; }

        public FixedValueFallback(object value) => Value = value;

        public PropertyMappingResult GetProperty(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, int columnIndex, string stringValue)
        {
            return PropertyMappingResult.Success(Value);
        }
    }
}

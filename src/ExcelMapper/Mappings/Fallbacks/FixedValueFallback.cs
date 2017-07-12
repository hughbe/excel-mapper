namespace ExcelMapper.Mappings.Fallbacks
{
    public class FixedValueFallback : IFallbackItem
    {
        public object Value { get; }

        public FixedValueFallback(object value) => Value = value;        

        public object PerformFallback(ExcelSheet sheet, int rowIndex, ReadResult result)
        {
            return Value;
        }
    }
}

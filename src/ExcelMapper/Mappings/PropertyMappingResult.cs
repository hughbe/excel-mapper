namespace ExcelMapper.Mappings
{
    public enum PropertyMappingResultType
    {
        Success,
        Invalid
    }

    public struct PropertyMappingResult
    {
        public static PropertyMappingResult Success(object value) => new PropertyMappingResult
        {
            Value = value,
            Type = PropertyMappingResultType.Success
        };

        public static PropertyMappingResult Invalid() => new PropertyMappingResult
        {
            Type = PropertyMappingResultType.Invalid
        };

        public object Value { get; private set; }
        public PropertyMappingResultType Type { get; private set; }
    }
}

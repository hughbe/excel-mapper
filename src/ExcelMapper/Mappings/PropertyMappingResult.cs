namespace ExcelMapper.Mappings
{
    public enum PropertyMappingResultType
    {
        Began,
        Continue,
        Success,
        Invalid
    }

    public struct PropertyMappingResult
    {
        public static PropertyMappingResult Began(object value) => new PropertyMappingResult
        {
            Value = value,
            Type = PropertyMappingResultType.Began
        };

        public static PropertyMappingResult Success(object value) => new PropertyMappingResult
        {
            Value = value,
            Type = PropertyMappingResultType.Success
        };

        public static PropertyMappingResult Continue() => new PropertyMappingResult
        {
            Type = PropertyMappingResultType.Continue
        };

        public static PropertyMappingResult Invalid() => new PropertyMappingResult
        {
            Type = PropertyMappingResultType.Invalid
        };

        public object Value { get; private set; }
        public PropertyMappingResultType Type { get; private set; }
    }
}

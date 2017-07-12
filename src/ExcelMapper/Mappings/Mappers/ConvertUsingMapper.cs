using System;

namespace ExcelMapper.Mappings.Mappers
{
    public delegate PropertyMappingResultType ConvertUsingMappingDelegate(ReadResult readResult, ref object value);

    public class ConvertUsingMapper : IStringValueMapper
    {
        public ConvertUsingMappingDelegate Converter { get; }

        public ConvertUsingMapper(ConvertUsingMappingDelegate converter)
        {
            Converter = converter ?? throw new ArgumentNullException(nameof(converter));
        }

        public PropertyMappingResultType GetProperty(ReadResult readResult, ref object value)
        {
            return Converter(readResult, ref value);
        }
    }
}

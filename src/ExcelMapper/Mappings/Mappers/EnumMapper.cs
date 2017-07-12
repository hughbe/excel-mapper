using System;
using System.Reflection;

namespace ExcelMapper.Mappings.Mappers
{
    public class EnumMapper : IStringValueMapper
    {
        public Type EnumType { get; }

        public EnumMapper(Type enumType)
        {
            if (enumType == null)
            {
                throw new ArgumentNullException(nameof(enumType));
            }

            if (!enumType.GetTypeInfo().IsEnum)
            {
                throw new ArgumentException($"Type {enumType} is not an Enum.", nameof(enumType));
            }

            EnumType = enumType;
        }

        public PropertyMappingResultType GetProperty(ReadResult readResult, ref object value)
        {
            try
            {
                value = Enum.Parse(EnumType, readResult.StringValue);
                return PropertyMappingResultType.Success;
            }
            catch
            {
                return PropertyMappingResultType.Invalid;
            }
        }
    }
}

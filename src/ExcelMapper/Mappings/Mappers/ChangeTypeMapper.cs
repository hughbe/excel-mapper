using System;
using ExcelMapper.Utilities;

namespace ExcelMapper.Mappings.Mappers
{
    public class ChangeTypeMapper : IStringValueMapper
    {
        public Type Type { get; }

        public ChangeTypeMapper(Type type)
        {
            if (type == null)
            {
                throw new ArgumentNullException(nameof(type));
            }

            if (!type.ImplementsInterface(typeof(IConvertible)))
            {
                throw new ArgumentException($"Type \"{type}\" must implement IConvertible to support Convert.ChangeType.", nameof(type));
            }

            Type = type;
        }

        public PropertyMappingResultType GetProperty(ReadResult readResult, ref object value)
        {
            try
            {
                value = Convert.ChangeType(readResult.StringValue, Type);
                return PropertyMappingResultType.Success;
            }
            catch
            {
                return PropertyMappingResultType.Invalid;
            }
        }
    }
}

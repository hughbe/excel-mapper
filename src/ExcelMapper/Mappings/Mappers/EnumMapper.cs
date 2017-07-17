using System;
using System.Reflection;

namespace ExcelMapper.Mappings.Mappers
{
    /// <summary>
    /// A mapper that tries to map the value of a cell to an enum of a given type.
    /// </summary>
    public class EnumMapper : ICellValueMapper
    {
        /// <summary>
        /// Gets the type of the enum to map the value of a cell to.
        /// </summary>
        public Type EnumType { get; }

        /// <summary>
        /// Constructs a mapper that tries to map the value of a cell to an enum of a given type.
        /// </summary>
        /// <param name="enumType">The type of the enum to convert the value of a cell to.</param>
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

        public PropertyMappingResultType GetProperty(ReadCellValueResult readResult, ref object value)
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

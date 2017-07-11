using System;
using System.Reflection;
using ExcelDataReader;

namespace ExcelMapper.Mappings.Items
{
    public class ParseAsEnumMappingItem : ISinglePropertyMappingItem
    {
        public Type EnumType { get; }

        public ParseAsEnumMappingItem(Type enumType)
        {
            if (enumType == null)
            {
                throw new ArgumentNullException(nameof(enumType));
            }

            if (!enumType.GetTypeInfo().IsEnum)
            {
                throw new ArgumentException(nameof(enumType), $"Type {enumType} is not an Enum.");
            }

            EnumType = enumType;
        }

        public PropertyMappingResult GetProperty(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, int columnIndex, string stringValue)
        {
            try
            {
                object value = Enum.Parse(EnumType, stringValue);
                return PropertyMappingResult.Success(value);
            }
            catch
            {
                return PropertyMappingResult.Invalid();
            }
        }
    }
}

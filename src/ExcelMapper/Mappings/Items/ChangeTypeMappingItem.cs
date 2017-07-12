using System;
using ExcelDataReader;
using ExcelMapper.Utilities;

namespace ExcelMapper.Mappings.Items
{
    public class ChangeTypeMappingItem : ISinglePropertyMappingItem
    {
        public Type Type { get; }

        public ChangeTypeMappingItem(Type type)
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

        public PropertyMappingResult GetProperty(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MapResult mapResult)
        {
            try
            {
                object value = Convert.ChangeType(mapResult.StringValue, Type);
                return PropertyMappingResult.Success(value);
            }
            catch
            {
                return PropertyMappingResult.Invalid();
            }
        }
    }
}

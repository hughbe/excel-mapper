using System;
using ExcelDataReader;

namespace ExcelMapper.Mappings.Items
{
    internal class ChangeTypeMappingItem : ISinglePropertyMappingItem
    {
        public Type Type { get; }

        public ChangeTypeMappingItem(Type type)
        {
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

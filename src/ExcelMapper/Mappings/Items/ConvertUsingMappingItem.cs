using System;
using ExcelDataReader;

namespace ExcelMapper.Mappings.Items
{
    public delegate PropertyMappingResult ConvertUsingMappingDelegate(MapResult mapResult);

    internal class ConvertUsingMappingItem : ISinglePropertyMappingItem
    {
        public ConvertUsingMappingDelegate Converter { get; }

        public ConvertUsingMappingItem(ConvertUsingMappingDelegate converter)
        {
            Converter = converter ?? throw new ArgumentNullException(nameof(converter));
        }

        public PropertyMappingResult GetProperty(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, MapResult mapResult)
        {
            return Converter(mapResult);
        }
    }
}

using System;
using ExcelDataReader;

namespace ExcelMapper.Mappings.Items
{
    public delegate PropertyMappingResult ConvertUsingMappingDelegate(ReadResult mapResult);

    public class ConvertUsingMappingItem : ISinglePropertyMappingItem
    {
        public ConvertUsingMappingDelegate Converter { get; }

        public ConvertUsingMappingItem(ConvertUsingMappingDelegate converter)
        {
            Converter = converter ?? throw new ArgumentNullException(nameof(converter));
        }

        public PropertyMappingResult GetProperty(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, ReadResult mapResult)
        {
            return Converter(mapResult);
        }
    }
}

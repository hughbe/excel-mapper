using System;
using ExcelDataReader;

namespace ExcelMapper.Mappings
{
    internal class OptionalMapping : ISinglePropertyMapper
    {
        public ISinglePropertyMapper Mapper { get; set; }

        public OptionalMapping(ISinglePropertyMapper mapper)
        {
            Mapper = mapper ?? throw new ArgumentNullException(nameof(mapper));
        }

        public MapResult GetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader)
        {
            try
            {
                return Mapper.GetValue(sheet, rowIndex, reader);
            }
            catch
            {
                return new MapResult(0, null);
            }
        }
    }
}

using System;
using ExcelDataReader;

namespace ExcelMapper.Mappings
{
    public class OptionalPropertyMapper : ISinglePropertyMapper
    {
        public ISinglePropertyMapper _mapper;

        public ISinglePropertyMapper Mapper
        {
            get => _mapper;
            set => _mapper = value ?? throw new ArgumentNullException(nameof(value));
        }

        public OptionalPropertyMapper(ISinglePropertyMapper mapper)
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

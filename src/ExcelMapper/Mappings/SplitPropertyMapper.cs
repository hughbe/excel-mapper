using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDataReader;

namespace ExcelMapper.Mappings
{
    public class SplitPropertyMapper : IMultiPropertyMapper
    {
        private char[] _separators = new char[] { ',' };

        public char[] Separators
        {
            get => _separators;
            set
            {
                if (value == null)
                {
                    throw new ArgumentNullException(nameof(value));
                }

                if (value.Length == 0)
                {
                    throw new ArgumentException("Separators cannot be empty.", nameof(value));
                }

                _separators = value;
            }
        }

        public StringSplitOptions Options { get; set; }

        private ISinglePropertyMapper _mapper;
        public ISinglePropertyMapper Mapper
        {
            get => _mapper;
            set => _mapper = value ?? throw new ArgumentNullException(nameof(value));
        }

        public SplitPropertyMapper(ISinglePropertyMapper mapper)
        {
            Mapper = mapper ?? throw new ArgumentNullException(nameof(mapper));
        }

        public IEnumerable<MapResult> GetValues(ExcelSheet sheet, int rowIndex, IExcelDataReader reader)
        {
            MapResult mapResult = Mapper.GetValue(sheet, rowIndex, reader);
            if (mapResult.StringValue == null)
            {
                return Enumerable.Empty<MapResult>();
            }

            string[] splitStringValues = mapResult.StringValue.Split(Separators, Options);
            return splitStringValues.Select(s => new MapResult(mapResult.ColumnIndex, s));
        }
    }
}

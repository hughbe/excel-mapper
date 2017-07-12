using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDataReader;

namespace ExcelMapper.Mappings.Readers
{
    public class SplitColumnReader : IMultipleValuesReader
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

        private ISingleValueReader _columnReader;
        public ISingleValueReader ColumnReader
        {
            get => _columnReader;
            set => _columnReader = value ?? throw new ArgumentNullException(nameof(value));
        }

        public SplitColumnReader(ISingleValueReader columnReader)
        {
            ColumnReader = columnReader ?? throw new ArgumentNullException(nameof(columnReader));
        }

        public IEnumerable<ReadResult> GetValues(ExcelSheet sheet, int rowIndex, IExcelDataReader reader)
        {
            ReadResult mapResult = ColumnReader.GetValue(sheet, rowIndex, reader);
            if (mapResult.StringValue == null)
            {
                return Enumerable.Empty<ReadResult>();
            }

            string[] splitStringValues = mapResult.StringValue.Split(Separators, Options);
            return splitStringValues.Select(s => new ReadResult(mapResult.ColumnIndex, s));
        }
    }
}

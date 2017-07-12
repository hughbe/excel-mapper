using System;
using ExcelDataReader;

namespace ExcelMapper.Mappings.Readers
{
    public class OptionalColumnReader : ISingleValueReader
    {
        public ISingleValueReader _innerReader;

        public ISingleValueReader InnerReader
        {
            get => _innerReader;
            set => _innerReader = value ?? throw new ArgumentNullException(nameof(value));
        }

        public OptionalColumnReader(ISingleValueReader innerReader)
        {
            InnerReader = innerReader ?? throw new ArgumentNullException(nameof(innerReader));
        }

        public ReadResult GetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader)
        {
            try
            {
                return InnerReader.GetValue(sheet, rowIndex, reader);
            }
            catch
            {
                return new ReadResult(0, null);
            }
        }
    }
}

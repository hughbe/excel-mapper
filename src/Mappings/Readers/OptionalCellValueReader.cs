using System;
using ExcelDataReader;

namespace ExcelMapper.Mappings.Readers
{
    /// <summary>
    /// Reads the value of a cell and 
    /// </summary>
    public class OptionalCellValueReader : ICellValueReader
    {
        private ICellValueReader _innerReader;

        public ICellValueReader InnerReader
        {
            get => _innerReader;
            set => _innerReader = value ?? throw new ArgumentNullException(nameof(value));
        }

        public OptionalCellValueReader(ICellValueReader innerReader)
        {
            InnerReader = innerReader ?? throw new ArgumentNullException(nameof(innerReader));
        }

        public ReadCellValueResult GetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader)
        {
            try
            {
                return InnerReader.GetValue(sheet, rowIndex, reader);
            }
            catch
            {
                return new ReadCellValueResult(0, null);
            }
        }
    }
}

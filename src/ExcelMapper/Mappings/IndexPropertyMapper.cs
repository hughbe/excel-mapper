using System;
using ExcelDataReader;

namespace ExcelMapper.Mappings
{
    public sealed class IndexPropertyMapper : ISinglePropertyMapper
    {
        public int ColumnIndex { get; }

        public IndexPropertyMapper(int columnIndex)
        {
            if (columnIndex < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(columnIndex), columnIndex, $"Column index {columnIndex} must be greater or equal to zero.");
            }

            ColumnIndex = columnIndex;
        }

        public MapResult GetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader)
        {
            return new MapResult(ColumnIndex, reader.GetString(ColumnIndex));
        }
    }
}

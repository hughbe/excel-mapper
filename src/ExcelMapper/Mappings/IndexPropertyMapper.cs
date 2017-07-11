using System;
using ExcelDataReader;

namespace ExcelMapper.Mappings
{
    internal sealed class IndexPropertyMapper : ISinglePropertyMapper
    {
        public int Index { get; }

        public IndexPropertyMapper(int index)
        {
            if (index < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            Index = index;
        }

        public int GetColumnIndex(ExcelSheet sheet, int rowIndex, IExcelDataReader reader)
        {
            return Index;
        }
    }
}

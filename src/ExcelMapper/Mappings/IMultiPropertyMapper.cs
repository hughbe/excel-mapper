using System.Collections.Generic;
using ExcelDataReader;

namespace ExcelMapper.Mappings
{
    public interface IMultiPropertyMapper
    {
        int CapacityEstimate { get; }
        IEnumerable<int> GetColumnIndices(ExcelSheet sheet, int rowIndex, IExcelDataReader reader);
    }
}

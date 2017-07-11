using System.Collections.Generic;
using ExcelDataReader;

namespace ExcelMapper.Mappings
{
    public interface IMultiPropertyMapper
    {
        IEnumerable<MapResult> GetValues(ExcelSheet sheet, int rowIndex, IExcelDataReader reader);
    }
}

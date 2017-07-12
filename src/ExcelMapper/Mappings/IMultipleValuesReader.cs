using System.Collections.Generic;
using ExcelDataReader;

namespace ExcelMapper.Mappings
{
    public interface IMultipleValuesReader
    {
        IEnumerable<ReadResult> GetValues(ExcelSheet sheet, int rowIndex, IExcelDataReader reader);
    }
}

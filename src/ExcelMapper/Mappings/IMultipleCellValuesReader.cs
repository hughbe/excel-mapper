using System.Collections.Generic;
using ExcelDataReader;

namespace ExcelMapper.Mappings
{
    /// <summary>
    /// An interface that describes an object that can read the values of multiple cells in a row.
    /// This describes a one-to-many or many-to-many mapping between a cell and a mapped value.
    /// </summary>
    public interface IMultipleCellValuesReader
    {
        IEnumerable<ReadCellValueResult> GetValues(ExcelSheet sheet, int rowIndex, IExcelDataReader reader);
    }
}

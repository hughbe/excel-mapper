using System.Collections.Generic;
using ExcelDataReader;

namespace ExcelMapper.Abstractions
{
    /// <summary>
    /// An interface that describes an object that can read the values of multiple cells in a row.
    /// This describes a one-to-many or many-to-many mapping between a cell and a mapped value.
    /// </summary>
    public interface IMultipleCellValuesReader
    {
        bool TryGetValues(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, out IEnumerable<ReadCellValueResult> result);
    }
}

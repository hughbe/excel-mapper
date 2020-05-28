using System.Collections.Generic;
using System.Linq;
using ExcelDataReader;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Readers
{
    /// <summary>
    /// Reads a multiple values of all columns in a sheet.
    /// </summary>
    public sealed class AllColumnNamesValueReader : IMultipleCellValuesReader
    {
        public bool TryGetValues(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, out IEnumerable<ReadCellValueResult> result)
        {
            if (sheet.Heading == null)
            {
                throw new ExcelMappingException($"The sheet \"{sheet.Name}\" does not have a heading. Use a column index mapping instead.");
            }

            result = sheet.Heading.ColumnNames.Select(columnName =>
            {
                var index = sheet.Heading.GetColumnIndex(columnName);
                var value = reader[index]?.ToString();
                return new ReadCellValueResult(index, value);
            });
            return true;
        }
    }
}

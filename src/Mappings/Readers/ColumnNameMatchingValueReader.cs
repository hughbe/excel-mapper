using System;
using ExcelDataReader;

namespace ExcelMapper.Mappings.Readers
{
    /// <summary>
    /// Reads the value of a single cell given the predicate matching the column name.
    /// </summary>
    public sealed class ColumnNameMatchingValueReader : ICellValueReader
    {
        private readonly Func<string, bool> _predicate;

        /// <summary>
        /// Constructs a reader that reads the value of a single cell given the predicate matching the column name.
        /// </summary>
        /// <param name="predicate">The predicate containing the column name to read.</param>
        public ColumnNameMatchingValueReader(Func<string, bool> predicate)
        {
            _predicate = predicate ?? throw new ArgumentNullException(nameof(predicate));
        }

        public ReadCellValueResult GetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader)
        {
            if (sheet.Heading == null)
            {
                throw new ExcelMappingException($"The sheet \"{sheet.Name}\" does not have a heading. Use a column index mapping instead.");
            }

            int index = sheet.Heading.GetFirstColumnMatchingIndex(_predicate);
            string value = reader[index]?.ToString();
            return new ReadCellValueResult(index, value);
        }
    }
}
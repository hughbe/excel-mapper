using System;
using ExcelDataReader;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Readers
{
    /// <summary>
    /// Reads the value of a single cell given the predicate matching the column name.
    /// </summary>
    public sealed class ColumnNameMatchingValueReader : ISingleCellValueReader
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

        public bool TryGetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader, out ReadCellValueResult result)
        {
            if (sheet.Heading == null)
            {
                throw new ExcelMappingException($"The sheet \"{sheet.Name}\" does not have a heading. Use a column index mapping instead.");
            }

            if (!sheet.Heading.TryGetFirstColumnMatchingIndex(_predicate, out int index))
            {
                result = default;
                return false;
            }

            string value = reader[index]?.ToString();
            result = new ReadCellValueResult(index, value);
            return true;
        }
    }
}
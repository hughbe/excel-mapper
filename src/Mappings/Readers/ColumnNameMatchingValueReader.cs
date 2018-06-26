namespace ExcelMapper.Mappings.Readers
{
    using System;

    using ExcelDataReader;

    /// <summary>
    /// Reads the value of a single cell given the predicate matching the column name.
    /// </summary>
    public sealed class ColumnNameMatchingValueReader : ICellValueReader
    {
        private readonly Func<string, bool> predicate;

        /// <summary>
        /// Constructs a reader that reads the value of a single cell given the predicate matching the column name.
        /// </summary>
        /// <param name="predicate">The predicate containing the column name to read.</param>
        public ColumnNameMatchingValueReader(Func<string, bool> predicate)
        {
            if (predicate == null)
            {
                throw new ArgumentNullException(nameof(predicate));
            }

            this.predicate = predicate;
        }

        public ReadCellValueResult GetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader)
        {
            if (sheet.Heading == null)
            {
                throw new ExcelMappingException($"The sheet \"{sheet.Name}\" does not have a heading. Use a column index mapping instead.");
            }

            var index = sheet.Heading.GetFirstColumnMatchingIndex(this.predicate);
            var value = reader[index]?.ToString();
            return new ReadCellValueResult(index, value);
        }
    }
}
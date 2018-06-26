namespace ExcelMapper.Mappings.Readers
{
    using System;

    using ExcelDataReader;

    public sealed class ColumnNameMatchingValueReader : ICellValueReader
    {
        private readonly Func<string, bool> predicate;

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
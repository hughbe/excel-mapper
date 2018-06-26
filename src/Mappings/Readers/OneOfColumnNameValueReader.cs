namespace ExcelMapper.Mappings.Readers
{
    using System;
    using System.Linq;

    using ExcelDataReader;

    using ExcelMapper.Utilities;

    public sealed class OneOfColumnNameValueReader : ICellValueReader
    {
        public OneOfColumnNameValueReader(string[] columnNames)
        {
            if (columnNames == null)
            {
                throw new ArgumentNullException(nameof(columnNames));
            }

            if (columnNames.Length == 0)
            {
                throw new ArgumentException("Column names cannot be empty.", nameof(columnNames));
            }

            if (columnNames.Any(columnName => columnName == null))
            {
                throw new ArgumentException($"Null column name in {columnNames.ArrayJoin()}.", nameof(columnNames));
            }

            this.ColumnNames = columnNames;
        }

        public string[] ColumnNames { get; set; }

        public ReadCellValueResult GetValue(ExcelSheet sheet, int rowIndex, IExcelDataReader reader)
        {
            if (sheet.Heading == null)
            {
                throw new ExcelMappingException($"The sheet \"{sheet.Name}\" does not have a heading. Use a column index mapping instead.");
            }

            var index = sheet.Heading.GetFirstColumnIndex(this.ColumnNames);
            var value = reader[index]?.ToString();
            return new ReadCellValueResult(index, value);
        }
    }
}
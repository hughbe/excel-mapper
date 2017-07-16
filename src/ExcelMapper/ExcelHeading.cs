using System.Collections.Generic;
using System.Linq;
using ExcelDataReader;

namespace ExcelMapper
{
    public class ExcelHeading
    {
        internal ExcelHeading(IExcelDataReader reader)
        {
            var nameMapping = new Dictionary<string, int>(reader.FieldCount);
            var columnNames = new string[reader.FieldCount];

            for (int columnIndex = 0; columnIndex < reader.FieldCount; columnIndex++)
            {
                string columnName = reader.GetString(columnIndex);
                if (columnName == null)
                {
                    columnNames[columnIndex] = string.Empty;

                }
                else
                {
                    nameMapping.Add(columnName, columnIndex);
                    columnNames[columnIndex] = columnName;
                }
            }

            NameMapping = nameMapping;
            _columnNames = columnNames;
        }

        private string[] _columnNames{ get; }
        private Dictionary<string, int> NameMapping { get; }

        public string GetColumnName(int index) => _columnNames[index];

        public int GetColumnIndex(string columnName)
        {
            if (!NameMapping.TryGetValue(columnName, out int index))
            {
                string foundColumns = string.Join(", ", NameMapping.Keys.Select(c => $"\"{c}\""));
                throw new ExcelMappingException($"Column \"{columnName}\" does not exist in [{foundColumns}]");
            }

            return index;
        }

        public IEnumerable<string> ColumnNames => _columnNames;
    }
}

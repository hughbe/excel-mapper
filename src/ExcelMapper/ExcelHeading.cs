using System.Collections.Generic;
using System.Linq;
using ExcelDataReader;

namespace ExcelMapper
{
    public class ExcelHeading
    {
        internal ExcelHeading(IExcelDataReader reader)
        {
            var indexMapping = new Dictionary<int, string>(reader.FieldCount);
            var nameMapping = new Dictionary<string, int>(reader.FieldCount);
            for (int columnIndex = 0; columnIndex < reader.FieldCount; columnIndex++)
            {
                string columnName = reader.GetString(columnIndex);
                indexMapping.Add(columnIndex, columnName);
                nameMapping.Add(columnName, columnIndex);
            }

            IndexMapping = indexMapping;
            NameMapping = nameMapping;
        }

        private Dictionary<int, string> IndexMapping { get; }
        private Dictionary<string, int> NameMapping { get; }

        public string GetColumnName(int index) => IndexMapping[index];

        public int GetColumnIndex(string columnName)
        {
            if (!NameMapping.TryGetValue(columnName, out int index))
            {
                string foundColumns = string.Join(", ", NameMapping.Keys.Select(c => $"\"{c}\""));
                throw new ExcelMappingException($"Column \"{columnName}\" does not exist in [{foundColumns}]");
            }

            return index;
        }

        public IEnumerable<string> ColumnNames => NameMapping.Keys;
    }
}

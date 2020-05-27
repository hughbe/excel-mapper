using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace ExcelMapper
{
    /// <summary>
    /// An object that represents the heading read from a sheet.
    /// </summary>
    public class ExcelHeading
    {
        private readonly string[] _columnNames;

        internal ExcelHeading(IDataRecord reader)
        {
            var nameMapping = new Dictionary<string, int>(reader.FieldCount, StringComparer.OrdinalIgnoreCase);
            var columnNames = new string[reader.FieldCount];

            for (int columnIndex = 0; columnIndex < reader.FieldCount; columnIndex++)
            {
                string columnName = reader.GetValue(columnIndex)?.ToString();
                if (columnName == null)
                {
                    columnNames[columnIndex] = string.Empty;

                }
                else
                {
                    if (nameMapping.ContainsKey(columnName))
                    {
                        columnName += "_" + Guid.NewGuid();
                    }
                    nameMapping.Add(columnName, columnIndex);
                    columnNames[columnIndex] = columnName;
                }
            }

            NameMapping = nameMapping;
            _columnNames = columnNames;
        }

        private Dictionary<string, int> NameMapping { get; }

        /// <summary>
        /// Gets the name of the column at the given zero-based index.
        /// </summary>
        /// <param name="index">The zero-based index to get the name of.</param>
        /// <returns>The name of the column at the given zero-based index.</returns>
        public string GetColumnName(int index) => _columnNames[index];

        /// <summary>
        /// Gets the zero-based index of the column with the given name. This method throws an ExcelMappingException
        /// if the column does not exist.
        /// </summary>
        /// <param name="columnName">The name of the column to get the zero-based index of.</param>
        /// <returns>The zero-based index of the column with the given name.</returns>
        public int GetColumnIndex(string columnName)
        {
            if (!NameMapping.TryGetValue(columnName, out int index))
            {
                string foundColumns = string.Join(", ", NameMapping.Keys.Select(c => $"\"{c}\""));
                throw new ExcelMappingException($"Column \"{columnName}\" does not exist in [{foundColumns}]");
            }

            return index;
        }

        /// <summary>
        /// Tries to get the zero-based index of the column with the given name.
        /// </summary>
        /// <param name="columnName">The name of the column to get the zero-based index of.</param>
        /// <param name="index">The zero-based index of the column with the given name.</param>
        /// <returns>Whether or not the column exists.</returns>
        public bool TryGetColumnIndex(string columnName, out int index) => NameMapping.TryGetValue(columnName, out index);

        /// <summary>
        /// Gets the zero-based index of the column with the given name if it matches the supplied predicate.
        /// This method throws an ExcelMappingException if the column does not exist.
        /// </summary>
        /// <param name="predicate">The predicate containing the names of the column to get the zero-based index of.</param>
        /// <returns>The zero-based index of the column with the given name.</returns>
        public int GetFirstColumnMatchingIndex(Func<string, bool> predicate)
        {
            var key = NameMapping.Keys.FirstOrDefault(predicate);

            if (key == null)
            {
                string foundColumns = string.Join(", ", NameMapping.Keys.Select(c => $"\"{c}\""));
                throw new ExcelMappingException($"No Columns found matching predicate from [{foundColumns}]");
            }

            return NameMapping[key];
        }

        /// <summary>
        /// Tries to get the zero-based index of the column with the given name if it matches the supplied predicate.
        /// </summary>
        /// <param name="predicate">The predicate containing the names of the column to get the zero-based index of.</param>
        /// <param name="index">The zero-based index of the column with the given name.</param>
        /// <returns>Whether or not the column exists.</returns>
        public bool TryGetFirstColumnMatchingIndex(Func<string, bool> predicate, out int index)
        {
            string key = NameMapping.Keys.FirstOrDefault(predicate);
            if (key == null)
            {
                index = -1;
                return false;
            }

            index = NameMapping[key];
            return true;
        }

        /// <summary>
        /// Gets the list of all column names in the heading.
        /// </summary>
        public IEnumerable<string> ColumnNames => _columnNames;
    }
}

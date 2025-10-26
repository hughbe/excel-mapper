﻿using System.Data;
using System.Linq;

namespace ExcelMapper;

/// <summary>
/// An object that represents the heading read from a sheet.
/// </summary>
public class ExcelHeading
{
    private readonly string[] _columnNames;

    internal ExcelHeading(IDataRecord reader, ExcelRange dataRange, ExcelImporterConfiguration configuration)
    {
        if (reader.FieldCount > configuration.MaxColumnsPerSheet)
        {
            throw new ExcelMappingException(
                $"Sheet has {reader.FieldCount} columns which exceeds the maximum allowed " +
                $"({configuration.MaxColumnsPerSheet}). Increase MaxColumnsPerSheet in the configuration if this is a legitimate file.");
        }

        var (offset, length) = dataRange.Columns.GetOffsetAndLength(reader.FieldCount);

        var nameMapping = new SortedList<string, int>(length, StringComparer.OrdinalIgnoreCase);
        var columnNames = new string[length];

        for (var columnIndex = 0; columnIndex < length; columnIndex++)
        {
            var columnName = reader.GetValue(columnIndex + offset)?.ToString() ?? string.Empty;
            columnNames[columnIndex] = columnName; // Store original name
            
            // For mapping, use unique key if duplicate
            var mappingKey = columnName;
            if (nameMapping.ContainsKey(mappingKey))
            {
                // Use incremental counter for deduplication to ensure uniqueness
                int suffix = 2;
                do
                {
                    mappingKey = $"{columnName}_{suffix}";
                    suffix++;
                } while (nameMapping.ContainsKey(mappingKey));
            }

            nameMapping.Add(mappingKey, columnIndex);
        }

        NameMapping = nameMapping;
        _columnNames = columnNames;
    }

    private SortedList<string, int> NameMapping { get; }

    /// <summary>
    /// Gets the name of the column at the given zero-based index.
    /// </summary>
    /// <param name="columnIndex">The zero-based index to get the name of.</param>
    /// <returns>The name of the column at the given zero-based index.</returns>
    public string GetColumnName(int columnIndex)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(columnIndex);
        ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(columnIndex, _columnNames.Length);
        return _columnNames[columnIndex];
    }

    /// <summary>
    /// Gets the zero-based index of the column with the given name. This method throws an ExcelMappingException
    /// if the column does not exist.
    /// </summary>
    /// <param name="columnName">The name of the column to get the zero-based index of.</param>
    /// <returns>The zero-based index of the column with the given name.</returns>
    public int GetColumnIndex(string columnName)
    {
        ArgumentNullException.ThrowIfNull(columnName);

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
    public bool TryGetColumnIndex(string columnName, out int index)
    {
        ArgumentNullException.ThrowIfNull(columnName);
        return NameMapping.TryGetValue(columnName, out index);
    }

    /// <summary>
    /// Gets the zero-based index of the column with the given name if it matches the supplied predicate.
    /// This method throws an ExcelMappingException if the column does not exist.
    /// </summary>
    /// <param name="predicate">The predicate containing the names of the column to get the zero-based index of.</param>
    /// <returns>The zero-based index of the column with the given name.</returns>
    public int GetFirstColumnMatchingIndex(Func<string, bool> predicate)
    {
        ArgumentNullException.ThrowIfNull(predicate);
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
        ArgumentNullException.ThrowIfNull(predicate);

        string? key = NameMapping.Keys.FirstOrDefault(predicate);
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
    public IReadOnlyList<string> ColumnNames => _columnNames;
}

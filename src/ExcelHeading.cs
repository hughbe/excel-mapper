using System.Data;
using System.Diagnostics.CodeAnalysis;
using System.Linq;

namespace ExcelMapper;

/// <summary>
/// An object that represents the heading read from a sheet.
/// </summary>
public sealed class ExcelHeading
{
    internal ExcelHeading(IDataRecord reader, ExcelRange dataRange, int maxColumns)
    {
        if (reader.FieldCount > maxColumns)
        {
            throw new ExcelMappingException(
                $"Sheet has {reader.FieldCount} columns which exceeds the maximum allowed " +
                $"({maxColumns}). Increase MaxColumnsPerSheet in the configuration if this is a legitimate file.");
        }

        var (offset, length) = dataRange.Columns.GetOffsetAndLength(reader.FieldCount);

        var nameMapping = new Dictionary<string, int>(length, StringComparer.OrdinalIgnoreCase);
        var columnNames = new string[length];

        for (var columnIndex = 0; columnIndex < length; columnIndex++)
        {
            var columnName = reader.GetValue(columnIndex + offset)?.ToString() ?? string.Empty;
            columnNames[columnIndex] = columnName;

            // A name that appears more than once maps to its first column. Mapping a member to it
            // is an error, as it isn't clear which column to read (see HasDuplicateNames).
            if (nameMapping.ContainsKey(columnName))
            {
                // Blank heading cells aren't names: members can't be mapped to them.
                HasDuplicateNames |= columnName.Length != 0;
                continue;
            }

            // Store the index of the column in the sheet (not relative to the data range) so
            // that it can be passed directly to the underlying reader.
            nameMapping.Add(columnName, columnIndex + offset);
        }

        ColumnOffset = offset;
        ColumnNames = columnNames;
        NameMapping = nameMapping;
    }

    /// <summary>
    /// Gets the zero-based index in the sheet of the first column in the heading. This is
    /// non-zero if the data range of the sheet does not start at the first column.
    /// </summary>
    internal int ColumnOffset { get; }

    /// <summary>
    /// Gets the list of all column names in the heading.
    /// </summary>
    public IReadOnlyList<string> ColumnNames { get; }

    private Dictionary<string, int> NameMapping { get; }

    /// <summary>
    /// Whether two columns have the same name, ignoring case, so looking up a column by name must
    /// check that it's the only one.
    /// </summary>
    internal bool HasDuplicateNames { get; }

    /// <summary>
    /// Gets the name of the column at the given zero-based index in the sheet.
    /// </summary>
    /// <param name="columnIndex">The zero-based index in the sheet of the column to get the name of.
    /// If the sheet's data range does not start at the first column, this is not the same as the
    /// index into <see cref="ColumnNames"/>.</param>
    /// <returns>The name of the column at the given zero-based index.</returns>
    public string GetColumnName(int columnIndex)
    {
        ThrowHelpers.ThrowIfLessThan(columnIndex, ColumnOffset, nameof(columnIndex));
        ThrowHelpers.ThrowIfGreaterThanOrEqual(columnIndex, ColumnOffset + ColumnNames.Count, nameof(columnIndex));
        return ColumnNames[columnIndex - ColumnOffset];
    }

    internal bool TryGetColumnName(int columnIndex, [NotNullWhen(true)] out string? columnName)
    {
        var relativeIndex = columnIndex - ColumnOffset;
        if (relativeIndex < 0 || relativeIndex >= ColumnNames.Count)
        {
            columnName = null;
            return false;
        }

        columnName = ColumnNames[relativeIndex];
        return true;
    }

    /// <summary>
    /// Gets the zero-based index in the sheet of the column with the given name. This method throws an ExcelMappingException
    /// if the column does not exist. If more than one column has the name, this is the first.
    /// </summary>
    /// <param name="columnName">The name of the column to get the zero-based index of.</param>
    /// <returns>The zero-based index of the column with the given name.</returns>
    public int GetColumnIndex(string columnName)
    {
        ThrowHelpers.ThrowIfNull(columnName, nameof(columnName));

        if (!NameMapping.TryGetValue(columnName, out int index))
        {
            var foundColumns = string.Join(", ", ColumnNames.Select(static c => $"\"{c}\""));
            throw new ExcelMappingException($"Column \"{columnName}\" does not exist in [{foundColumns}]");
        }

        return index;
    }

    /// <summary>
    /// Tries to get the zero-based index in the sheet of the column with the given name.
    /// </summary>
    /// <param name="columnName">The name of the column to get the zero-based index of.</param>
    /// <param name="index">The zero-based index of the column with the given name. If more than one
    /// column has the name, this is the first.</param>
    /// <returns>Whether or not the column exists.</returns>
    public bool TryGetColumnIndex(string columnName, out int index)
    {
        ThrowHelpers.ThrowIfNull(columnName, nameof(columnName));
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
        if (!TryGetFirstColumnMatchingIndex(predicate, out var index))
        {
            var foundColumns = string.Join(", ", ColumnNames.Select(static c => $"\"{c}\""));
            throw new ExcelMappingException($"No Columns found matching predicate from [{foundColumns}]");
        }

        return index;
    }

    /// <summary>
    /// Tries to get the zero-based index of the column with the given name if it matches the supplied predicate.
    /// </summary>
    /// <param name="predicate">The predicate containing the names of the column to get the zero-based index of.</param>
    /// <param name="index">The zero-based index of the column with the given name.</param>
    /// <returns>Whether or not the column exists.</returns>
    public bool TryGetFirstColumnMatchingIndex(Func<string, bool> predicate, out int index)
    {
        ThrowHelpers.ThrowIfNull(predicate, nameof(predicate));

        // Search in column order.
        for (var i = 0; i < ColumnNames.Count; i++)
        {
            if (predicate(ColumnNames[i]))
            {
                index = i + ColumnOffset;
                return true;
            }
        }

        index = -1;
        return false;
    }
}

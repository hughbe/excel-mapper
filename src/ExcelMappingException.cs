using System.Linq;
using System.Reflection;

namespace ExcelMapper;

/// <summary>
/// Represents errors that occur during Excel mapping.
/// </summary>
public class ExcelMappingException : Exception
{
    /// <summary>
    /// Creates an ExcelMappingException with the default message.
    /// </summary>
    public ExcelMappingException()
    {
    }

    /// <summary>
    /// Creates an ExcelMappingException with the given message.
    /// </summary>
    /// <param name="message">The message of the exception.</param>
    public ExcelMappingException(string message) : base(message)
    {
    }

    /// <summary>
    /// Creates an ExcelMappingException with the given message and inner exception.
    /// </summary>
    /// <param name="message">The message of the exception.</param>
    /// <param name="innerException">The inner exception of the exception.</param>
    public ExcelMappingException(string message, Exception innerException) : base(message, innerException)
    {
    }

    /// <summary>
    /// Creates an ExcelMappingException throw trying to map a cell value to a property or field.
    /// </summary>
    /// <param name="message">The base error message of the exception.</param>
    /// <param name="sheet">The sheet that is currently being read.</param>
    /// <param name="rowIndex">The zero-based index of the row in the sheet that is currently being read.</param>
    /// <param name="columnIndex">The zero-based index of the column in the sheet that is currently being read.</param>
    public ExcelMappingException(string message, ExcelSheet sheet, int rowIndex, int columnIndex) : this(message, sheet, rowIndex, columnIndex, null)
    {
    }

    /// <summary>
    /// Creates an ExcelMappingException throw trying to map a cell value to a property or field.
    /// </summary>
    /// <param name="message">The base error message of the exception.</param>
    /// <param name="sheet">The sheet that is currently being read.</param>
    /// <param name="rowIndex">The zero-based index of the row in the sheet that is currently being read.</param>
    /// <param name="columnIndex">The zero-based index of the column in the sheet that is currently being read.</param>
    /// <param name="innerException">The inner exception of the exception.</param>
    public ExcelMappingException(string message, ExcelSheet sheet, int rowIndex, int columnIndex, Exception? innerException)
        : this(message, sheet, rowIndex, columnIndex, innerException, hint: null)
    {
    }

    private ExcelMappingException(string message, ExcelSheet sheet, int rowIndex, int columnIndex, Exception? innerException, string? hint)
        : base(GetMessage(message, sheet, rowIndex, columnIndex) + hint, innerException)
    {
        Sheet = sheet;
        RowIndex = rowIndex;
        ColumnIndex = columnIndex;
    }

    /// <summary>
    /// The sheet that is currently being read.
    /// </summary>
    public ExcelSheet? Sheet { get; }

    /// <summary>
    /// The zero-based index of the row in the sheet that is currently being read.
    /// </summary>
    public int RowIndex { get; } = -1;

    /// <summary>
    /// The zero-based index of the column in the sheet that is currently being read.
    /// </summary>
    public int ColumnIndex { get; } = -1;

    /// <summary>
    /// The Excel address (for example "B11") of the cell that was being read, or null if the
    /// row or column is not known.
    /// </summary>
    public string? CellAddress => RowIndex >= 0 && ColumnIndex >= 0 ? ExcelRange.GetCellAddress(RowIndex, ColumnIndex) : null;

    internal static ExcelMappingException CreateForNoSuchColumn(ExcelSheet sheet, int rowIndex, object readerFactory, MemberInfo? member)
    {
        if (readerFactory is IColumnNameProviderCellReaderFactory columnNameProvider)
        {
            var columnName = columnNameProvider.GetColumnName(sheet);
            if (!string.IsNullOrEmpty(columnName))
            {
                return new ExcelMappingException($"Could not read value for member \"{member?.Name}\" for column \"{columnName}\"", sheet, rowIndex, -1, null, GetSimilarColumnHint(sheet, columnName!));
            }
        }
        if (readerFactory is IColumnIndexProviderCellReaderFactory columnIndexProvider)
        {
            var columnIndex = columnIndexProvider.GetColumnIndex(sheet);
            if (columnIndex != null && columnIndex != -1)
            {
                return new ExcelMappingException($"Could not read value for member \"{member?.Name}\" for column at index {columnIndex}", sheet, rowIndex, -1);
            }
        }
        if (readerFactory is IColumnNamesProviderCellReaderFactory columnNamesProvider)
        {
            if (columnNamesProvider.GetColumnNames(sheet) is { } columnNames)
            {
                return CreateForMultipleColumnNames(sheet, rowIndex, member, columnNames);
            }
        }
        if (readerFactory is IColumnIndicesProviderCellReaderFactory columnIndicesProvider)
        {
            if (columnIndicesProvider.GetColumnIndices(sheet) is { } columnIndices)
            {
                return CreateForMultipleColumnIndices(sheet, rowIndex, member, columnIndices);
            }
        }

        return new ExcelMappingException($"Could not read value for member \"{member?.Name}\"", sheet, rowIndex, -1);
    }

    /// <summary>
    /// Suggests a column whose name only differs by case or by separators (spaces, underscores,
    /// hyphens and dots), which is a common reason for a column not being found.
    /// </summary>
    private static string? GetSimilarColumnHint(ExcelSheet sheet, string columnName)
    {
        if (sheet.Heading is null || sheet.Configuration.IgnoreSeparatorsInColumnNames)
        {
            return null;
        }

        var normalizedName = ColumnUtilities.RemoveSeparators(columnName);
        foreach (var candidate in sheet.Heading.ColumnNames)
        {
            if (!string.Equals(candidate, columnName, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(ColumnUtilities.RemoveSeparators(candidate), normalizedName, StringComparison.OrdinalIgnoreCase))
            {
                return $" Did you mean column \"{candidate}\"? Use ExcelColumnName or WithColumnName, or set IgnoreSeparatorsInColumnNames in the importer's configuration.";
            }
        }

        return null;
    }

    private static string GetMessage(string message, ExcelSheet? sheet, int rowIndex, int columnIndex)
    {
        string columnDescription = string.Empty;
        if (columnIndex >= 0)
        {
            // The column may lie outside the heading's data range, in which case only report
            // its cell address rather than throwing while building the message.
            if (sheet != null && sheet.HasHeading
                && (sheet.Heading ?? sheet.ReadHeading()).TryGetColumnName(columnIndex, out var columnName))
            {
                columnDescription = $" in column \"{columnName}\"";
            }
        }

        // Report the location as Excel shows it: a cell address such as "B11" if the column is
        // known, otherwise a one-based row number.
        string location;
        if (rowIndex >= 0 && columnIndex >= 0)
        {
            location = $" at cell {ExcelRange.GetCellAddress(rowIndex, columnIndex)}";
        }
        else if (rowIndex >= 0)
        {
            location = $" on row {rowIndex + 1}";
        }
        else if (columnIndex >= 0)
        {
            location = $" in column {ExcelRange.GetColumnLetters(columnIndex)}";
        }
        else
        {
            location = string.Empty;
        }

        return $"{message}{columnDescription}{location} in sheet \"{sheet?.Name}\".";
    }

    private static ExcelMappingException CreateForMultipleColumnNames(ExcelSheet sheet, int rowIndex, MemberInfo? member, IReadOnlyList<string> columnNames)
    {
        if (columnNames.Count == 0)
        {
            return new ExcelMappingException($"Could not read value for member \"{member?.Name}\" (no columns matching)", sheet, rowIndex, -1);
        }
        if (columnNames.Count == 1)
        {
            return new ExcelMappingException($"Could not read value for member \"{member?.Name}\" for column \"{columnNames[0]}\"", sheet, rowIndex, -1);
        }

        return new ExcelMappingException($"Could not read value for member \"{member?.Name}\" for columns {string.Join(", ", columnNames.Select(c => $"\"{c}\""))}", sheet, rowIndex, -1);
    }

    private static ExcelMappingException CreateForMultipleColumnIndices(ExcelSheet sheet, int rowIndex, MemberInfo? member, IReadOnlyList<int> columnIndices)
    {
        if (columnIndices.Count == 0)
        {
            return new ExcelMappingException($"Could not read value for member \"{member?.Name}\" (no columns matching)", sheet, rowIndex, -1);
        }
        if (columnIndices.Count == 1)
        {
            return new ExcelMappingException($"Could not read value for member \"{member?.Name}\" for column at index {columnIndices[0]}", sheet, rowIndex, -1);
        }

        return new ExcelMappingException($"Could not read value for member \"{member?.Name}\" for columns at indices {string.Join(", ", columnIndices)}", sheet, rowIndex, -1);
    }
}

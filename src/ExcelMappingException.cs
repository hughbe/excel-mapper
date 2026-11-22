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

    /// <summary>
    /// Creates an ExcelMappingException for the cell a converter, transformer or fallback is given,
    /// for example for a value it can't convert. The message gives the cell's address and column
    /// name, also when reading a region.
    /// </summary>
    /// <param name="message">The base error message of the exception.</param>
    /// <param name="context">The cell.</param>
    public ExcelMappingException(string message, CellContext context) : this(message, context, null)
    {
    }

    /// <summary>
    /// Creates an ExcelMappingException for the cell a converter, transformer or fallback is given,
    /// for example for a value it can't convert. The message gives the cell's address and column
    /// name, also when reading a region.
    /// </summary>
    /// <param name="message">The base error message of the exception.</param>
    /// <param name="context">The cell.</param>
    /// <param name="innerException">The inner exception of the exception.</param>
    public ExcelMappingException(string message, CellContext context, Exception? innerException)
        : this(message, GetReadingSheet(context), context.RowIndex, context.ColumnIndex, innerException, hint: null)
    {
        Member = context.Member;
    }

    private static ExcelSheet GetReadingSheet(CellContext context)
        => context.ReadingSheet ?? throw new ArgumentException("The context has no sheet.", nameof(context));

    internal ExcelMappingException(string message, ExcelSheet sheet, int rowIndex, int columnIndex, Exception? innerException, string? hint)
        : base(GetMessage(message, sheet, rowIndex, columnIndex) + hint, innerException)
    {
        // The message and column name use the heading of the rows being read (a region's, for
        // the sheet a region is read with), but the sheet is the one the user has.
        Sheet = sheet?.UserSheet;
        RowIndex = rowIndex;
        ColumnIndex = columnIndex;
        if (columnIndex >= 0 && sheet?.Heading is { } heading && heading.TryGetColumnName(columnIndex, out var columnName))
        {
            ColumnName = columnName;
        }
    }

    /// <summary>
    /// The name of the column in the heading that was being read, or the name of a column that
    /// could not be found. Null if the column or the sheet's heading is not known.
    /// </summary>
    public string? ColumnName { get; internal set; }

    /// <summary>
    /// The property or field that was being mapped, if known.
    /// </summary>
    public MemberInfo? Member { get; internal set; }

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

    internal static ExcelMappingException CreateForNoSuchColumn(ExcelSheet sheet, int rowIndex, Readers.Columns columns, MemberInfo? member)
    {
        var exception = columns.CreateMissingException(sheet, rowIndex, member);
        exception.Member = member;
        exception.ColumnName ??= columns.ColumnName;
        return exception;
    }

    /// <summary>
    /// Suggests a column whose name only differs by case or by separators (spaces, underscores,
    /// hyphens and dots), which is a common reason for a column not being found.
    /// </summary>
    internal static string? GetSimilarColumnHint(ExcelSheet sheet, string columnName)
    {
        if (sheet.Heading is null || sheet.Settings.IgnoreSeparatorsInColumnNames)
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
            // Only use a heading that has already been read: building an error message must not
            // move the reader. The column may also lie outside the heading's data range, in which
            // case only its cell address is reported.
            if (sheet?.Heading is { } heading && heading.TryGetColumnName(columnIndex, out var columnName))
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
}

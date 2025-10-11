using System;
using System.Linq;
using System.Reflection;

namespace ExcelMapper;

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
        : base(GetMessage(message, sheet, rowIndex, columnIndex), innerException)
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
    
    internal static ExcelMappingException CreateForNoSuchColumn(ExcelSheet sheet, int rowIndex, object readerFactory, MemberInfo? member)
    {
        if (readerFactory is IColumnNameProviderCellReaderFactory columnNameProvider)
        {
            var columnName = columnNameProvider.GetColumnName(sheet);
            throw new ExcelMappingException($"Could not read value for member \"{member?.Name}\" for column \"{columnName}\"", sheet, rowIndex, -1);
        }
        if (readerFactory is IColumnIndexProviderCellReaderFactory columnIndexProvider)
        {
            var columnIndex = columnIndexProvider.GetColumnIndex(sheet);
            throw new ExcelMappingException($"Could not read value for member \"{member?.Name}\" for column at index {columnIndex}", sheet, rowIndex, -1);
        }
        if (readerFactory is IColumnNamesProviderCellReaderFactory columnNamesProvider && columnNamesProvider.GetColumnNames(sheet) is { } columnNames)
        {
            if (columnNames.Length == 0)
            {
                return new ExcelMappingException($"Could not read value for member \"{member?.Name}\" (no columns matching)", sheet, rowIndex, -1);
            }
            if (columnNames.Length == 1)
            {
                return new ExcelMappingException($"Could not read value for member \"{member?.Name}\" for column \"{columnNames[0]}\"", sheet, rowIndex, -1);
            }

            return new ExcelMappingException($"Could not read value for member \"{member?.Name}\" for columns {string.Join(", ", columnNames.Select(c => $"\"{c}\""))}", sheet, rowIndex, -1);
        }
        if (readerFactory is IColumnIndicesProviderCellReaderFactory columnIndicesProvider && columnIndicesProvider.GetColumnIndices(sheet) is { } columnIndices)
        {
            if (columnIndices.Length == 1)
            {
                return new ExcelMappingException($"Could not read value for member \"{member?.Name}\" for column at index {columnIndices[0]}", sheet, rowIndex, -1);
            }

            return new ExcelMappingException($"Could not read value for member \"{member?.Name}\" for columns at indices {string.Join(", ", columnIndices)}", sheet, rowIndex, -1);
        }

        return new ExcelMappingException($"Could not read value for member \"{member?.Name}\"", sheet, rowIndex, -1);
    }

    private static string GetMessage(string message, ExcelSheet? sheet, int rowIndex, int columnIndex)
    {
        string position = string.Empty;
        if (columnIndex != -1)
        {
            if (sheet != null && sheet.HasHeading)
            {
                var heading = sheet.Heading ?? sheet.ReadHeading();

                position = $" in column \"{heading.GetColumnName(columnIndex)}\"";
            }
            else
            {
                position = $" in position \"{columnIndex}\"";
            }
        }

        return $"{message}{position} on row {rowIndex} in sheet \"{sheet?.Name}\".";
    }
}

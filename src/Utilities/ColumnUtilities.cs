using System.Runtime.CompilerServices;

namespace ExcelMapper.Utilities;

internal static class ColumnUtilities
{
    public static void ValidateColumnIndex(int columnIndex, [CallerArgumentExpression(nameof(columnIndex))] string? paramName = null)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(columnIndex, paramName);
    }

    public static void ValidateColumnIndices(IList<int> columnIndices, [CallerArgumentExpression(nameof(columnIndices))] string? paramName = null)
    {
        ArgumentNullException.ThrowIfNull(columnIndices, paramName);
        if (columnIndices.Count == 0)
        {
            throw new ArgumentException("Column indices cannot be empty.", paramName);
        }


        foreach (var columnIndex in columnIndices)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(columnIndex, paramName);
        }
    }

    public static void ValidateColumnName(string columnName, [CallerArgumentExpression(nameof(columnName))] string? paramName = null)
    {
        ArgumentNullException.ThrowIfNull(columnName, paramName);
        ArgumentException.ThrowIfNullOrEmpty(columnName, paramName);
    }

    public static void ValidateColumnNames(IList<string> columnNames, [CallerArgumentExpression(nameof(columnNames))] string? paramName = null)
    {
        ArgumentNullException.ThrowIfNull(columnNames, paramName);
        if (columnNames.Count == 0)
        {
            throw new ArgumentException("Column names cannot be empty.", paramName);
        }

        foreach (string columnName in columnNames)
        {
            if (columnName == null)
            {
                throw new ArgumentException("Column names cannot contain null values.", paramName);
            }

            ArgumentException.ThrowIfNullOrEmpty(columnName, paramName);
        }
    }
}

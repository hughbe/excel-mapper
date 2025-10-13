using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelMapper.Utilities;

internal static class ColumnUtilities
{
    public static void ValidateColumnIndex(int columnIndex, string paramName)
    {
        if (columnIndex < 0)
        {
            throw new ArgumentOutOfRangeException(paramName, columnIndex, $"Column {columnIndex} must be greater or equal to zero.");
        }
    }

    public static void ValidateColumnIndices(IList<int> columnIndices, string paramName)
    {
        if (columnIndices == null)
        {
            throw new ArgumentNullException(paramName);
        }

        if (columnIndices.Count == 0)
        {
            throw new ArgumentException("Column indices cannot be empty.", paramName);
        }


        foreach (var columnIndex in columnIndices)
        {
            if (columnIndex < 0)
            {
                throw new ArgumentOutOfRangeException(paramName, columnIndex, $"Column {columnIndex} must be greater or equal to zero.");
            }
        }
    }

    public static void ValidateColumnName(string columnName, string paramName)
    {
        if (columnName == null)
        {
            throw new ArgumentNullException(paramName);
        }

        if (columnName.Length == 0)
        {
            throw new ArgumentException("Column name cannot be empty.", paramName);
        }
    }

    public static void ValidateColumnNames(IList<string> columnNames, string paramName)
    {
        if (columnNames == null)
        {
            throw new ArgumentNullException(paramName);
        }
        if (columnNames.Count == 0)
        {
            throw new ArgumentException("Column names cannot be empty.", paramName);
        }

        foreach (string columnName in columnNames)
        {
            if (columnName == null)
            {
                throw new ArgumentException($"Null column name in {QuoteJoin(columnNames)}.", paramName);
            }
        }
    }

    private static string QuoteJoin(IEnumerable<string> values)
    {
        var quoted = values.Select(v => $"\"{v}\"");
        return $"[{string.Join(", ", quoted)}]";
    }
}

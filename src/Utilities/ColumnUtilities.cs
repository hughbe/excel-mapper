using System;
using System.Collections.Generic;

namespace ExcelMapper.Utilities;

internal static class ColumnUtilities
{
    public static void ValidateColumnIndex(int columnIndex, string paramName)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(columnIndex, paramName);
    }

    public static void ValidateColumnIndices(IList<int> columnIndices, string paramName)
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

    public static void ValidateColumnName(string columnName, string paramName)
    {
        ArgumentNullException.ThrowIfNull(columnName, paramName);
        ArgumentException.ThrowIfNullOrEmpty(columnName, paramName);
    }

    public static void ValidateColumnNames(IList<string> columnNames, string paramName)
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

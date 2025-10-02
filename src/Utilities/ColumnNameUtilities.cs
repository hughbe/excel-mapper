using System;

namespace ExcelMapper.Utilities;

internal static class ColumnNameUtilities
{
    public static void ValidateColumnNames(string[] columnNames, string paramName)
    {
        if (columnNames == null)
        {
            throw new ArgumentNullException(paramName);
        }
        if (columnNames.Length == 0)
        {
            throw new ArgumentException("Column names cannot be empty.", paramName);
        }

        foreach (string columnName in columnNames)
        {
            if (columnName == null)
            {
                throw new ArgumentException($"Null column name in {columnNames.ArrayJoin()}.", paramName);
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
}

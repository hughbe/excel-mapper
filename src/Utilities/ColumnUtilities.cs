using System.Runtime.CompilerServices;

namespace ExcelMapper.Utilities;

internal static class ColumnUtilities
{
    public static void ValidateColumnIndex(int columnIndex, [CallerArgumentExpression(nameof(columnIndex))] string? paramName = null)
    {
        ThrowHelpers.ThrowIfNegative(columnIndex, paramName);
    }

    public static void ValidateColumnIndices(IReadOnlyList<int> columnIndices, [CallerArgumentExpression(nameof(columnIndices))] string? paramName = null)
    {
        ThrowHelpers.ThrowIfNull(columnIndices, paramName);
        if (columnIndices.Count == 0)
        {
            throw new ArgumentException("Column indices cannot be empty.", paramName);
        }


        foreach (var columnIndex in columnIndices)
        {
            ThrowHelpers.ThrowIfNegative(columnIndex, paramName);
        }
    }

    public static void ValidateColumnName(string columnName, [CallerArgumentExpression(nameof(columnName))] string? paramName = null)
    {
        ThrowHelpers.ThrowIfNull(columnName, paramName);
        ThrowHelpers.ThrowIfNullOrEmpty(columnName, paramName);
    }

    public static void ValidateColumnNames(IReadOnlyList<string> columnNames, [CallerArgumentExpression(nameof(columnNames))] string? paramName = null)
    {
        ThrowHelpers.ThrowIfNull(columnNames, paramName);
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

            ThrowHelpers.ThrowIfNullOrEmpty(columnName, paramName);
        }
    }

    public static bool TryGetColumnIndex(ExcelSheet sheet, string columnName, StringComparison comparison, out int columnIndex)
    {
        ThrowHelpers.ThrowIfNull(sheet, nameof(sheet));
        if (sheet.Heading == null)
        {
            throw new ExcelMappingException($"The sheet \"{sheet.Name}\" does not have a heading. Use a column index mapping instead.");
        }

        if (TryGetExactColumnIndex(sheet, sheet.Heading, columnName, comparison, out columnIndex))
        {
            return true;
        }

        // Optionally fall back to ignoring separators, for example matching "First Name" to FirstName.
        if (sheet.Configuration.IgnoreSeparatorsInColumnNames)
        {
            var normalizedName = RemoveSeparators(columnName);
            var columnNames = sheet.Heading.ColumnNames;
            for (int i = 0; i < columnNames.Count; i++)
            {
                if (string.Equals(RemoveSeparators(columnNames[i]), normalizedName, comparison))
                {
                    columnIndex = i + sheet.Heading.ColumnOffset;
                    return true;
                }
            }
        }

        columnIndex = -1;
        return false;
    }

    private static bool IsSeparator(char c) => char.IsWhiteSpace(c) || c == '_' || c == '-' || c == '.';

    internal static string RemoveSeparators(string value)
    {
        var hasSeparator = false;
        foreach (var c in value)
        {
            if (IsSeparator(c))
            {
                hasSeparator = true;
                break;
            }
        }

        if (!hasSeparator)
        {
            return value;
        }

        var builder = new System.Text.StringBuilder(value.Length);
        foreach (var c in value)
        {
            if (!IsSeparator(c))
            {
                builder.Append(c);
            }
        }

        return builder.ToString();
    }

    private static bool TryGetExactColumnIndex(ExcelSheet sheet, ExcelHeading heading, string columnName, StringComparison comparison, out int columnIndex)
    {
        // ExcelSheet.TryGetColumnIndex defaults to OrdinalIgnoreCase.
        // Ordinal comparison can be optimised by ensuring the column name is in the correct case.
        if (comparison == StringComparison.OrdinalIgnoreCase || comparison == StringComparison.Ordinal)
        {
            if (!heading.TryGetColumnIndex(columnName, out columnIndex))
            {
                return false;
            }

            if (comparison != StringComparison.OrdinalIgnoreCase)
            {
                // Verify that the comparison matches.
                var actualName = heading.GetColumnName(columnIndex);
                if (!string.Equals(actualName, columnName, comparison))
                {
                    columnIndex = -1;
                    return false;
                }
            }

            return true;
        }
        else
        {
            var columnNames = heading.ColumnNames;
            for (int i = 0; i < columnNames.Count; i++)
            {
                if (string.Equals(columnNames[i], columnName, comparison))
                {
                    columnIndex = i + heading.ColumnOffset;
                    return true;
                }
            }
        }

        columnIndex = -1;
        return false;
    }

    /// <summary>
    /// Gets the zero-based start (inclusive) and end (exclusive) indices in the sheet of the
    /// columns that can be mapped. If the sheet has a heading, this is the columns in the
    /// heading (i.e. the columns in the data range). Otherwise this is every column.
    /// </summary>
    public static (int Start, int End) GetColumnRange(ExcelSheet sheet)
    {
        if (sheet.Heading is { } heading)
        {
            return (heading.ColumnOffset, heading.ColumnOffset + heading.ColumnNames.Count);
        }

        return (0, sheet.NumberOfColumns);
    }
}

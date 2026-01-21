using System.Text.RegularExpressions;

namespace ExcelMapper.Readers;

/// <summary>
/// Matches column names based on a regular expression.
/// </summary>
public class RegexColumnMatcher : IExcelColumnMatcher
{
    /// <summary>
    /// Gets the regular expression used to match column names.
    /// </summary>
    public Regex Regex { get; }

    /// <summary>
    /// Constructs a matcher that matches column names based on the given regular expression.
    /// </summary>
    /// <param name="regex">The regular expression used to match column names.</param>
    public RegexColumnMatcher(Regex regex)
    {
        ThrowHelpers.ThrowIfNull(regex, nameof(regex));

        Regex = regex;
    }

    /// <inheritdoc/>
    public bool ColumnMatches(ExcelSheet sheet, int columnIndex)
    {
        ThrowHelpers.ThrowIfNull(sheet, nameof(sheet));
        if (sheet.Heading == null)
        {
            throw new ExcelMappingException($"The sheet \"{sheet.Name}\" does not have a heading. Use a column index mapping instead.");
        }

        var columnName = sheet.Heading.GetColumnName(columnIndex);
        return Regex.IsMatch(columnName);
    }
}
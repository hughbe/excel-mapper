using System.Text.RegularExpressions;

namespace ExcelMapper.Readers;

public class RegexColumnMatcher : IExcelColumnMatcher
{
    public Regex Regex { get; }

    public RegexColumnMatcher(Regex regex)
    {
        ArgumentNullException.ThrowIfNull(regex);

        Regex = regex;
    }

    public bool ColumnMatches(ExcelSheet sheet, int columnIndex)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        if (sheet.Heading == null)
        {
            throw new ExcelMappingException($"The sheet \"{sheet.Name}\" does not have a heading. Use a column index mapping instead.");
        }

        var columnName = sheet.Heading.GetColumnName(columnIndex);
        return Regex.IsMatch(columnName);
    }
}
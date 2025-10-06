using System;
using System.Text.RegularExpressions;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Readers;

public class RegexColumnMatcher : IExcelColumnMatcher
{
    public Regex Regex { get; }

    public RegexColumnMatcher(Regex regex)
    {
        if (regex == null)
        {
            throw new ArgumentNullException(nameof(regex));
        }

        Regex = regex;
    }

    public bool ColumnMatches(ExcelSheet sheet, int columnIndex)
    {
        if (sheet == null)
        {
            throw new ArgumentNullException(nameof(sheet));
        }
        if (sheet.Heading == null)
        {
            throw new ExcelMappingException($"The sheet \"{sheet.Name}\" does not have a heading. Use a column index mapping instead.");
        }

        var columnName = sheet.Heading.GetColumnName(columnIndex);
        return Regex.IsMatch(columnName);
    }
}
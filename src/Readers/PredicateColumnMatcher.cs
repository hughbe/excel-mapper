using System;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Readers;

public class PredicateColumnMatcher : IExcelColumnMatcher
{
    public Func<string, bool> Predicate { get; }

    public PredicateColumnMatcher(Func<string, bool> predicate)
    {
        ArgumentNullException.ThrowIfNull(predicate);
        Predicate = predicate;
    }

    public bool ColumnMatches(ExcelSheet sheet, int columnIndex)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        if (sheet.Heading == null)
        {
            throw new ExcelMappingException($"The sheet \"{sheet.Name}\" does not have a heading. Use a column index mapping instead.");
        }

        return Predicate(sheet.Heading.GetColumnName(columnIndex));
    }
}
using System;
using ExcelMapper.Abstractions;

namespace ExcelMapper.Readers;

public class PredicateColumnMatcher : IExcelColumnMatcher
{
    public Func<string, bool> Predicate { get; }

    public PredicateColumnMatcher(Func<string, bool> predicate)
    {
        Predicate = predicate ?? throw new ArgumentNullException(nameof(predicate));
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

        return Predicate(sheet.Heading.GetColumnName(columnIndex));
    }
}
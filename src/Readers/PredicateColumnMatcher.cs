namespace ExcelMapper.Readers;

/// <summary>
/// Matches column names based on a predicate.
/// </summary>
public class PredicateColumnMatcher : IExcelColumnMatcher
{
    /// <summary>
    /// Gets the predicate used to match column names.
    /// </summary>
    public Func<string, bool> Predicate { get; }

    /// <summary>
    /// Constructs a matcher that matches column names based on the given predicate.
    /// </summary>
    /// <param name="predicate">The predicate used to match column names.</param>
    public PredicateColumnMatcher(Func<string, bool> predicate)
    {
        ThrowHelpers.ThrowIfNull(predicate, nameof(predicate));
        Predicate = predicate;
    }

    /// <inheritdoc/>
    public bool ColumnMatches(ExcelSheet sheet, int columnIndex)
    {
        ThrowHelpers.ThrowIfNull(sheet, nameof(sheet));
        if (sheet.Heading == null)
        {
            throw new ExcelMappingException($"The sheet \"{sheet.Name}\" does not have a heading. Use a column index mapping instead.");
        }

        return Predicate(sheet.Heading.GetColumnName(columnIndex));
    }
}
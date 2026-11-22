namespace ExcelMapper;

/// <summary>
/// Chooses the columns a member is read from. Set it with <c>WithColumns(selector)</c>,
/// <see cref="ExcelColumnMatchingAttribute"/> or <see cref="ExcelColumnsMatchingAttribute"/>.
/// </summary>
/// <remarks>
/// The columns are chosen once for each sheet, not for every row. A member with a single value is
/// read from the first column chosen; a collection or dictionary is read from all of them. A
/// selector can be used for many sheets at once on different threads, so it mustn't change its own
/// state.
/// </remarks>
public interface IColumnSelector
{
    /// <summary>
    /// Chooses the columns to read.
    /// </summary>
    /// <param name="context">The sheet being read. Its <see cref="ColumnSelectorContext.Heading"/>
    /// gives the columns' names, if the rows have a heading.</param>
    /// <returns>The zero-based positions of the columns in the worksheet, as
    /// <see cref="ExcelHeading.GetColumnIndex(string)"/> gives them, in the order to read them, or
    /// an empty list if there are none. When reading a region, columns outside it are ignored.</returns>
    IReadOnlyList<int> SelectColumns(ColumnSelectorContext context);

    /// <summary>
    /// Describes the columns for error messages, for example <c>column "Name"</c> or
    /// <c>columns matching "^Q[1-4]$"</c>.
    /// </summary>
    /// <param name="context">The sheet being read.</param>
    string Describe(ColumnSelectorContext context);
}

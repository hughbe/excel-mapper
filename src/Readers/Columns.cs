using System.Linq;
using System.Reflection;

namespace ExcelMapper.Readers;

/// <summary>
/// The columns a member reads, chosen by name, position, a pattern or an <see cref="IColumnSelector"/>.
/// They're found once for each sheet (see <see cref="SheetView.FindColumn"/>): a member with a
/// single value reads the first column found, and a collection or dictionary reads all of them.
/// </summary>
/// <remarks>
/// Each kind also describes itself, for the error when its columns are missing and for
/// <see cref="ExcelSheet.ValidateHeading{T}()"/>.
/// </remarks>
internal abstract class Columns
{
    /// <summary>
    /// Finds the column a single value is read from.
    /// </summary>
    /// <returns>The column's position in the worksheet, or -1 if the sheet doesn't have it.</returns>
    public abstract int FindColumn(SheetView sheet);

    /// <summary>
    /// Finds the columns a collection or dictionary is read from.
    /// </summary>
    /// <returns>The columns' positions in the worksheet, in the order to read them, or null if
    /// the sheet doesn't have them.</returns>
    public virtual int[]? FindColumns(SheetView sheet)
    {
        var column = FindColumn(sheet);
        return column == -1 ? null : [column];
    }

    /// <summary>
    /// Whether the columns are looked up by name, so that they identify a heading row (see
    /// <see cref="ExcelSheet.FindHeading{T}()"/>).
    /// </summary>
    public virtual bool LooksUpNames => false;

    /// <summary>
    /// The column's name, for <see cref="ExcelMappingException.ColumnName"/>, if there's one.
    /// </summary>
    public virtual string? ColumnName => null;

    /// <summary>
    /// Creates the error for a member whose columns the sheet doesn't have.
    /// </summary>
    public virtual ExcelMappingException CreateMissingException(SheetView sheet, int rowIndex, MemberInfo? member)
        => Missing(sheet, member, "a matching column");

    /// <summary>
    /// Describes the columns for <see cref="ExcelSheet.ValidateHeading{T}()"/>, such as
    /// <c>column "Name"</c>.
    /// </summary>
    public virtual string Describe(SheetView sheet) => "a matching column";

    /// <summary>
    /// Creates the error for a member whose columns the sheet doesn't have, such as
    /// <c>Can't find column "Name" for member "Name" in sheet "People". The heading's columns are
    /// "First Name", "Age".</c>
    /// </summary>
    /// <param name="sheet">The sheet being read.</param>
    /// <param name="member">The member being read, if any.</param>
    /// <param name="columns">Describes the columns, such as <c>column "Name"</c>.</param>
    /// <param name="hint">Text to add after the message, such as a suggested column.</param>
    /// <param name="rowIndex">The row, if the columns are missing from it rather than the sheet.</param>
    private protected static ExcelMappingException Missing(SheetView sheet, MemberInfo? member, string columns, string? hint = null, int rowIndex = -1)
        => new($"Can't find {columns}{(member is null ? string.Empty : $" for member \"{member.Name}\"")}", sheet, rowIndex, -1, null, DescribeHeading(sheet) + hint);

    // The names in the heading, so a misspelt or renamed column can be seen in the message.
    private static string? DescribeHeading(SheetView sheet)
    {
        if (sheet.Heading is not { } heading)
        {
            return null;
        }

        const int MaxNames = 20;
        var names = heading.ColumnNames.Where(n => n.Length != 0).ToArray();
        if (names.Length == 0)
        {
            return " The heading is empty.";
        }

        var listed = string.Join(", ", names.Take(MaxNames).Select(n => $"\"{n}\""));
        var more = names.Length > MaxNames ? $" and {names.Length - MaxNames} more" : string.Empty;
        return $" The heading's columns are {listed}{more}.";
    }

    private protected static string DescribeNames(IReadOnlyList<string> names)
        => names.Count == 1 ? $"column \"{names[0]}\"" : $"columns {string.Join(", ", names.Select(n => $"\"{n}\""))}";
}

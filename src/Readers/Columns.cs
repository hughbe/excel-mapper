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
        => Missing(sheet, rowIndex, member, null);

    /// <summary>
    /// Describes the columns for <see cref="ExcelSheet.ValidateHeading{T}()"/>, such as
    /// <c>column "Name"</c>.
    /// </summary>
    public virtual string Describe(SheetView sheet) => "a matching column";

    private protected static ExcelMappingException Missing(SheetView sheet, int rowIndex, MemberInfo? member, string? columns)
        => new($"Could not read value for member \"{member?.Name}\"{(columns is null ? string.Empty : " for " + columns)}", sheet, rowIndex, -1);

    private protected static ExcelMappingException NoHeading(SheetView sheet)
        => new($"The sheet \"{sheet.Name}\" does not have a heading. Use a column index mapping instead.");

    private protected static string DescribeNames(IReadOnlyList<string> names)
        => names.Count == 1 ? $"column \"{names[0]}\"" : $"columns {string.Join(", ", names.Select(n => $"\"{n}\""))}";
}

using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

namespace ExcelMapper.Readers;

/// <summary>
/// The columns whose names match a predicate or a regular expression.
/// </summary>
internal sealed class MatchingColumns : Columns
{
    private readonly Func<string, bool> _matches;

    public MatchingColumns(Func<string, bool> predicate)
    {
        ThrowHelpers.ThrowIfNull(predicate, nameof(predicate));
        _matches = predicate;
    }

    public MatchingColumns(Regex regex)
    {
        ThrowHelpers.ThrowIfNull(regex, nameof(regex));
        _matches = regex.IsMatch;
    }

    public override int FindColumn(SheetView sheet)
    {
        var (start, end) = ColumnUtilities.GetColumnRange(sheet);
        for (var column = start; column < end; column++)
        {
            if (Matches(sheet, column))
            {
                return column;
            }
        }

        return -1;
    }

    public override int[]? FindColumns(SheetView sheet)
    {
        var columns = new List<int>();
        var (start, end) = ColumnUtilities.GetColumnRange(sheet);
        for (var column = start; column < end; column++)
        {
            if (Matches(sheet, column))
            {
                columns.Add(column);
            }
        }

        return columns.Count == 0 ? null : [.. columns];
    }

    private bool Matches(SheetView sheet, int column)
        => _matches((sheet.Heading ?? throw NoHeading(sheet)).GetColumnName(column));

    public override bool LooksUpNames => true;

    public override ExcelMappingException CreateMissingException(SheetView sheet, int rowIndex, MemberInfo? member)
    {
        if (sheet.Heading is not { } heading)
        {
            return Missing(sheet, rowIndex, member, null);
        }

        // The names of the matching columns, which a row may be too short to have.
        var names = new List<string>();
        var (start, end) = ColumnUtilities.GetColumnRange(sheet);
        for (var column = start; column < end; column++)
        {
            if (Matches(sheet, column))
            {
                names.Add(heading.GetColumnName(column)!);
            }
        }

        return names.Count == 0
            ? new($"Could not read value for member \"{member?.Name}\" (no columns matching)", sheet, rowIndex, -1)
            : Missing(sheet, rowIndex, member, DescribeNames(names));
    }
}

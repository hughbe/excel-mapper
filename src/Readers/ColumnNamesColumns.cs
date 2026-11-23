using System.Linq;
using System.Reflection;

namespace ExcelMapper.Readers;

/// <summary>
/// Columns with names: a single value reads the first the sheet has, and a collection or
/// dictionary reads all of them.
/// </summary>
internal sealed class ColumnNamesColumns : Columns
{
    public ColumnNamesColumns(IEnumerable<string> columnNames, StringComparison comparison = StringComparison.OrdinalIgnoreCase)
    {
        // Copied, so that changing the caller's array afterwards doesn't change the map.
        ThrowHelpers.ThrowIfNull(columnNames, nameof(columnNames));
        string[] names = [.. columnNames];
        ColumnUtilities.ValidateColumnNames(names, nameof(columnNames));
        EnumUtilities.ValidateIsDefined(comparison);
        Names = names;
        Comparison = comparison;
    }

    public IReadOnlyList<string> Names { get; }

    public StringComparison Comparison { get; }

    public override int FindColumn(SheetView sheet)
    {
        foreach (var name in Names)
        {
            if (ColumnUtilities.TryGetColumnIndex(sheet, name, Comparison, out var column))
            {
                return column;
            }
        }

        return -1;
    }

    public override int[]? FindColumns(SheetView sheet)
    {
        sheet.GetHeadingForNames();
        var columns = new int[Names.Count];
        for (int i = 0; i < columns.Length; i++)
        {
            if (!ColumnUtilities.TryGetColumnIndex(sheet, Names[i], Comparison, out columns[i]))
            {
                return null;
            }
        }

        return columns;
    }

    public override bool LooksUpNames => true;

    public override ExcelMappingException CreateMissingException(SheetView sheet, int rowIndex, MemberInfo? member)
        => Missing(sheet, member, DescribeNames(Names));

    public override string Describe(SheetView sheet) => $"one of columns {string.Join(", ", Names.Select(n => $"\"{n}\""))}";
}

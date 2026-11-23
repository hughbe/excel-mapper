using System.Linq;
using System.Reflection;

namespace ExcelMapper.Readers;

/// <summary>
/// Columns at positions: a single value reads the first the sheet has, and a collection or
/// dictionary reads all of them.
/// </summary>
internal sealed class ColumnIndicesColumns : Columns
{
    public ColumnIndicesColumns(IEnumerable<int> columnIndices)
    {
        // Copied, so that changing the caller's array afterwards doesn't change the map.
        ThrowHelpers.ThrowIfNull(columnIndices, nameof(columnIndices));
        int[] indices = [.. columnIndices];
        ColumnUtilities.ValidateColumnIndices(indices, nameof(columnIndices));
        Indices = indices;
    }

    public IReadOnlyList<int> Indices { get; }

    public override int FindColumn(SheetView sheet)
    {
        foreach (var index in Indices)
        {
            if (sheet.ToSheetColumn(index) is var column and not -1)
            {
                return column;
            }
        }

        return -1;
    }

    public override int[]? FindColumns(SheetView sheet)
    {
        var columns = new int[Indices.Count];
        for (int i = 0; i < columns.Length; i++)
        {
            columns[i] = sheet.ToSheetColumn(Indices[i]);
            if (columns[i] == -1)
            {
                return null;
            }
        }

        return columns;
    }

    public override ExcelMappingException CreateMissingException(SheetView sheet, int rowIndex, MemberInfo? member)
        => Missing(sheet, member, Describe(sheet));

    public override string Describe(SheetView sheet)
        => $"columns {string.Join(", ", Indices.Select(i => ExcelRange.GetColumnLetters(sheet.ColumnOffset + i)))}";
}

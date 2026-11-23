using System.Reflection;

namespace ExcelMapper.Readers;

/// <summary>
/// A column at a position, such as <c>[ExcelColumnIndex(2)]</c>. In a region, positions count from
/// the region's first column.
/// </summary>
internal sealed class ColumnIndexColumns : Columns
{
    public ColumnIndexColumns(int columnIndex)
    {
        ThrowHelpers.ThrowIfNegative(columnIndex, nameof(columnIndex));
        Index = columnIndex;
    }

    public int Index { get; }

    public override int FindColumn(SheetView sheet) => sheet.ToSheetColumn(Index);

    public override ExcelMappingException CreateMissingException(SheetView sheet, int rowIndex, MemberInfo? member)
        => Missing(sheet, member, Describe(sheet));

    public override string Describe(SheetView sheet)
        => $"column {ExcelRange.GetColumnLetters(sheet.ColumnOffset + Index)} (index {Index})";
}

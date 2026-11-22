using System.Linq;
using System.Reflection;

namespace ExcelMapper.Readers;

/// <summary>
/// One cell whose text is split into the values of a collection, such as "1,2,3": the cell is in
/// the first column <see cref="Cell"/> finds.
/// </summary>
internal sealed class SplitColumns(Columns cell, CellSplitter splitter) : Columns
{
    public Columns Cell { get; } = cell;

    public CellSplitter Splitter { get; } = splitter;

    public override int FindColumn(SheetView sheet) => Cell.FindColumn(sheet);

    public override bool LooksUpNames => Cell.LooksUpNames;

    public override ExcelMappingException CreateMissingException(SheetView sheet, int rowIndex, MemberInfo? member)
        => Cell switch
        {
            ColumnNameColumns or FirstColumnNameColumns => Missing(sheet, rowIndex, member, $"column \"{Cell.ColumnName}\""),
            ColumnIndexColumns index => Missing(sheet, rowIndex, member, $"column at index {index.Index}"),
            _ => Missing(sheet, rowIndex, member, null),
        };

    public override string Describe(SheetView sheet)
        => Cell switch
        {
            ColumnNameColumns or FirstColumnNameColumns => $"one of columns \"{Cell.ColumnName}\"",
            ColumnIndexColumns index => $"columns {ExcelRange.GetColumnLetters(sheet.ColumnOffset + index.Index)}",
            _ => base.Describe(sheet),
        };
}

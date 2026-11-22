using System.Linq;
using System.Reflection;

namespace ExcelMapper.Readers;

/// <summary>
/// The first of several columns that the sheet has, from several <see cref="ExcelColumnNameAttribute"/>s.
/// Missing columns are described by the first.
/// </summary>
internal sealed class FirstColumnNameColumns(IReadOnlyList<ColumnNameColumns> columns) : Columns
{
    public IReadOnlyList<ColumnNameColumns> Columns { get; } = columns;

    public override int FindColumn(SheetView sheet)
    {
        for (int i = 0; i < Columns.Count; i++)
        {
            if (Columns[i].FindColumn(sheet) is var column and not -1)
            {
                return column;
            }
        }

        return -1;
    }

    public override bool LooksUpNames => true;

    public override string ColumnName => Columns[0].Name;

    public override ExcelMappingException CreateMissingException(SheetView sheet, int rowIndex, MemberInfo? member)
        => Columns[0].CreateMissingException(sheet, rowIndex, member);

    public override string Describe(SheetView sheet) => Columns[0].Describe(sheet);
}

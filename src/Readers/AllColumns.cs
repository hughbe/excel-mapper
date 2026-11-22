using System.Linq;

namespace ExcelMapper.Readers;

/// <summary>
/// Every column with a name in the heading: the default for a dictionary of a row's columns.
/// </summary>
internal sealed class AllColumns : Columns
{
    public static AllColumns Instance { get; } = new();

    private AllColumns()
    {
    }

    public override int FindColumn(SheetView sheet) => throw new NotSupportedException();

    public override int[] FindColumns(SheetView sheet)
    {
        var heading = sheet.Heading ?? throw new ExcelMappingException($"The sheet \"{sheet.Name}\" does not have a heading. Use a column index map instead.");

        // Use the positions of the columns rather than looking them up by name, so that columns
        // with the same name are each read.
        var columns = new List<int>();
        var names = heading.ColumnNames;
        for (int i = 0; i < names.Count; i++)
        {
            if (!string.IsNullOrWhiteSpace(names[i]))
            {
                columns.Add(i + heading.ColumnOffset);
            }
        }

        return [.. columns];
    }
}

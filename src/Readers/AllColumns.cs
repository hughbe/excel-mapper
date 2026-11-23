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

    // Only collections and dictionaries read these columns, never a single value.
    public override int FindColumn(SheetView sheet) => throw new NotSupportedException();

    public override int[] FindColumns(SheetView sheet)
    {
        var heading = sheet.GetHeadingForAllColumns();

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

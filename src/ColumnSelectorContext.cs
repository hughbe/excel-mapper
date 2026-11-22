namespace ExcelMapper;

/// <summary>
/// What an <see cref="IColumnSelector"/> chooses columns from: the sheet being read and the heading
/// of its rows.
/// </summary>
public readonly struct ColumnSelectorContext
{
    /// <summary>
    /// Creates a context for a sheet, for example to test a selector.
    /// </summary>
    /// <param name="sheet">The sheet being read.</param>
    public ColumnSelectorContext(ExcelSheet sheet)
    {
        ThrowHelpers.ThrowIfNull(sheet, nameof(sheet));
        Sheet = sheet.UserSheet;
        Heading = sheet.Heading;
    }

    /// <summary>
    /// Gets the sheet being read. When reading a region, this is the sheet the region was created
    /// from.
    /// </summary>
    public ExcelSheet Sheet { get; }

    /// <summary>
    /// Gets the heading of the rows being read: the region's heading when reading a region,
    /// otherwise the sheet's. Null if they have no heading.
    /// </summary>
    public ExcelHeading? Heading { get; }

    /// <summary>
    /// Gets the number of columns in the worksheet.
    /// </summary>
    public int NumberOfColumns => Sheet.NumberOfColumns;
}

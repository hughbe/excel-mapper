namespace ExcelMapper.Utilities;

/// <summary>
/// Where a sheet, or a region of it, is reading: the row it last moved to. The reader is shared by
/// every sheet of the importer, so the cursor asks the <see cref="Utilities.RowSource"/> to move to
/// its row rather than moving the reader itself.
/// </summary>
internal sealed class RowCursor(Worksheet worksheet)
{
    /// <summary>
    /// Gets the index of the current row, or -1 before the first row.
    /// </summary>
    public int RowIndex { get; private set; } = -1;

    /// <summary>
    /// Moves back to before the first row: the next row read is the first one.
    /// </summary>
    public void Reset() => RowIndex = -1;

    /// <summary>
    /// Gets the number of rows in the worksheet.
    /// </summary>
    public int RowCount => worksheet.RowSource.GetRowCount(worksheet.Index);

    /// <summary>
    /// Gets the row index of a position that may count from the end of the worksheet.
    /// </summary>
    public int GetRowIndex(Index index) => index.IsFromEnd ? index.GetOffset(RowCount) : index.Value;

    /// <summary>
    /// Whether the current row's cells in the given columns are all empty.
    /// </summary>
    public bool IsBlankRow(int startColumnIndex, int endColumnIndex) => worksheet.RowSource.IsBlankRow(startColumnIndex, endColumnIndex);

    /// <summary>
    /// Moves to a row, throwing if the worksheet doesn't have it.
    /// </summary>
    public void MoveTo(int index)
    {
        if (!TryMoveTo(index))
        {
            throw new ExcelMappingException($"Sheet \"{worksheet.Name}\" does not have row {index}.");
        }
    }

    /// <summary>
    /// Moves to a row, so that its cells can be read. Other sheets of the importer share the
    /// reader, so this can mean reading forward or rewinding.
    /// </summary>
    /// <returns>False if the worksheet ends before the row: the cursor is then on its last row.</returns>
    public bool TryMoveTo(int index)
    {
        worksheet.ThrowIfClosed();

        var rowSource = worksheet.RowSource;
        var moved = rowSource.TryMoveTo(worksheet.Index, index);
        RowIndex = moved ? index : rowSource.LastRowIndex;

        // Reading rows of a sheet makes it the importer's current sheet, so ReadSheet() continues
        // from the sheet after it.
        worksheet.Importer.SheetIndex = worksheet.Index;
        return moved;
    }

    /// <summary>
    /// Moves to the next row of data in <paramref name="dataRange"/>, skipping rows before
    /// <paramref name="firstDataRow"/> and, if the importer skips blank lines, blank rows.
    /// </summary>
    /// <returns>False if there are no more rows.</returns>
    public bool MoveNext(ExcelRange dataRange, Index? firstDataRow)
    {
        worksheet.ThrowIfClosed();

        // Ensure we're on the correct sheet and row.
        MoveTo(RowIndex);

        // Check if we've reached the end of the data range. Only count the rows for ranges
        // relative to the end of the sheet: otherwise read until there are no more rows.
        var rows = dataRange.Rows;
        var endRowIndex = !rows.End.IsFromEnd ? rows.End.Value
            : rows.End.Value == 0 ? int.MaxValue
            : rows.End.GetOffset(RowCount);
        if (RowIndex + 1 >= endRowIndex)
        {
            return false;
        }

        // Skip any rows before the start of the data range. If the sheet has a heading, the
        // heading row is the start of the data range and has already been read.
        var startRowIndex = GetRowIndex(rows.Start);
        if (firstDataRow is { } first)
        {
            startRowIndex = Math.Max(startRowIndex, GetRowIndex(first));
        }

        if (RowIndex + 1 < startRowIndex)
        {
            if (startRowIndex >= endRowIndex || !TryMoveTo(startRowIndex - 1))
            {
                return false;
            }
        }

        // Read the next row.
        if (!TryMoveTo(RowIndex + 1))
        {
            return false;
        }

        if (worksheet.Configuration.SkipBlankLines)
        {
            // Only consider the columns in the data range when checking for blank rows.
            var fieldCount = worksheet.Reader.FieldCount;
            var startColumnIndex = Math.Min(dataRange.Columns.Start.GetOffset(fieldCount), fieldCount);
            var endColumnIndex = Math.Min(dataRange.Columns.End.GetOffset(fieldCount), fieldCount);

            while (IsBlankRow(startColumnIndex, endColumnIndex))
            {
                // Don't skip past the end of the data range.
                if (RowIndex + 1 >= endRowIndex || !TryMoveTo(RowIndex + 1))
                {
                    return false;
                }
            }
        }

        return true;
    }
}

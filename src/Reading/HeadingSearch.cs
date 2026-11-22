using System.Diagnostics.CodeAnalysis;

namespace ExcelMapper.Reading;

/// <summary>
/// Finds the heading of a sheet with rows above it, such as a title or notes: see
/// <see cref="ExcelSheet.FindHeading{T}()"/>.
/// </summary>
internal static class HeadingSearch
{
    /// <summary>
    /// The number of rows searched by default (documented as 50).
    /// </summary>
    public const int DefaultMaxRows = 50;

    /// <summary>
    /// Tries each row, from the sheet's heading index, as the heading: the first that has every
    /// column the class map needs by name and at least one of its columns is read as the heading.
    /// If none is, the sheet is unchanged and <paramref name="error"/> describes the closest row.
    /// </summary>
    public static bool TryFind<[DynamicallyAccessedMembers(TrimmingConstants.MappedType)] T>(SheetView sheet, RowCursor cursor, ExcelClassMap<T> classMap, int maxRows, [NotNullWhen(true)] out ExcelHeading? heading, [NotNullWhen(false)] out string? error)
    {
        var startRowIndex = cursor.GetRowIndex(sheet.DataRange.Rows.Start);
        var endRowIndex = (int)Math.Min((long)startRowIndex + maxRows, int.MaxValue);

        var closestRowIndex = -1;
        var closestFound = 0;
        List<string>? closestProblems = null;
        for (var rowIndex = startRowIndex; rowIndex < endRowIndex; rowIndex++)
        {
            if (!cursor.TryMoveTo(rowIndex))
            {
                break;
            }
            if (cursor.IsBlankRow(0, sheet.Worksheet.Reader.FieldCount))
            {
                continue;
            }

            // Check the class map's columns against this row as the heading.
            var candidate = new ExcelHeading(sheet.Worksheet.Reader, sheet.DataRange, sheet.Settings.MaxColumnsPerSheet);
            var lookedUp = 0;
            var found = 0;
            List<string>? problems = null;
            sheet.Heading = candidate;
            try
            {
                MapValidation.CountNamedColumns(classMap, sheet, ref lookedUp, ref found);
                if (found > 0)
                {
                    problems = [];
                    MapValidation.FindMissingColumns(classMap, sheet, null, problems);
                }
            }
            finally
            {
                sheet.Heading = null;
            }

            if (lookedUp == 0)
            {
                cursor.Reset();
                throw new ExcelMappingException($"Can't find the heading for type \"{typeof(T)}\" as it doesn't look up any column by name. Set {nameof(SheetView.HeadingIndex)} instead.");
            }

            if (found > 0 && problems!.Count == 0)
            {
                sheet.UseHeading(rowIndex, candidate);
                heading = candidate;
                error = null;
                return true;
            }

            if (found > closestFound)
            {
                closestRowIndex = rowIndex;
                closestFound = found;
                closestProblems = problems;
            }
        }

        // Leave the sheet as it was: the next read starts from the beginning of the sheet again.
        cursor.Reset();
        heading = null;
        var searched = $"the first {maxRows} rows{(startRowIndex == 0 ? string.Empty : $" from row {startRowIndex + 1}")} of sheet \"{sheet.Name}\"";
        error = closestRowIndex == -1
            ? $"Could not find the heading for type \"{typeof(T)}\": none of {searched} has any of its columns."
            : $"Could not find the heading for type \"{typeof(T)}\" in {searched}. The closest is row {closestRowIndex + 1}, which is missing {string.Join("; ", closestProblems!)}.";
        return false;
    }
}

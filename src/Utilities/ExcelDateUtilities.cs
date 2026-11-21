namespace ExcelMapper.Utilities;

/// <summary>
/// Helpers for Excel's serial date numbers: the number of days since 1899-12-30, with the time of
/// day as the fractional part.
/// </summary>
internal static class ExcelDateUtilities
{
    // The range accepted by DateTime.FromOADate.
    private const double MinSerialDate = -657435.0;
    private const double MaxSerialDate = 2958465.99999999;

    /// <summary>
    /// Converts a number from a cell without a date format, such as 45355 for 2024-03-04, to a date.
    /// </summary>
    public static bool TryGetDateTime(object? value, out DateTime result)
    {
        if (value is double serial && serial >= MinSerialDate && serial <= MaxSerialDate)
        {
            result = DateTime.FromOADate(serial);
            return true;
        }

        result = default;
        return false;
    }

    /// <summary>
    /// Excel stores a time without a date as a fraction of a day, which ExcelDataReader returns as a
    /// DateTime on Excel's first day (1899-12-30, or 1899-12-31 in some files).
    /// </summary>
    public static bool IsTimeOfDayOnly(DateTime value)
        => value.Date == new DateTime(1899, 12, 30) || value.Date == new DateTime(1899, 12, 31);
}

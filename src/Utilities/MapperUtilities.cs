namespace ExcelMapper.Utilities;

internal static class MapperUtilities
{
    /// <summary>
    /// Returns an invalid result whose exception is created only if it is needed, by running the
    /// throwing version of a parse that has already failed. This keeps the exact exception (for
    /// example FormatException or OverflowException) without throwing for every invalid cell.
    /// </summary>
    public static CellMapperResult InvalidFromParse(Action parse)
        => CellMapperResult.Invalid(() =>
        {
            try
            {
                parse();
            }
            catch (Exception exception)
            {
                return exception;
            }

            return new FormatException("The value could not be parsed.");
        });
}

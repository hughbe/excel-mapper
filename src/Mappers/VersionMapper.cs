namespace ExcelMapper.Mappers;

/// <summary>
/// Tries to map the value of a cell to <see cref="Version"/>.
/// </summary>
internal class VersionMapper : ICellMapper
{
    /// <inheritdoc/>
    public CellMapperResult Map(ReadCellResult readResult)
    {
        var stringValue = readResult.GetString();

        if (Version.TryParse(stringValue, out var version))
        {
            return CellMapperResult.Success(version);
        }

        // Only create the exception (by parsing again) if it is needed.
        return InvalidResult(stringValue);
    }

    // A separate method so that the closure capturing these values is only allocated for invalid
    // values. Creating the lambda in Map would allocate the closure for every cell.
    private static CellMapperResult InvalidResult(string? stringValue)
        => MapperUtilities.InvalidFromParse(() => new Version(stringValue!));
}

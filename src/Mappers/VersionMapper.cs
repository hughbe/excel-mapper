namespace ExcelMapper.Mappers;

/// <summary>
/// Tries to map the value of a cell to <see cref="Version"/>.
/// </summary>
public class VersionMapper : ICellMapper
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
        return MapperUtilities.InvalidFromParse(() => new Version(stringValue!));
    }
}

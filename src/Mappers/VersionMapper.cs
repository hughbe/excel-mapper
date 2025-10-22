namespace ExcelMapper.Mappers;

/// <summary>
/// Tries to map the value of a cell to <see cref="Version"/>.
/// </summary>
public class VersionMapper : ICellMapper
{
    public CellMapperResult Map(ReadCellResult readResult)
    {
        var stringValue = readResult.GetString();

        try
        {
            var version = new Version(stringValue!);
            return CellMapperResult.Success(version);
        }
        catch (Exception exception)
        {
            return CellMapperResult.Invalid(exception);
        }
    }
}

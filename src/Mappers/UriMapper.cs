namespace ExcelMapper.Mappers;

/// <summary>
/// Tries to map the value of a cell to an absolute <see cref="Uri"/>.
/// </summary>
public class UriMapper : ICellMapper
{
    public CellMapperResult Map(ReadCellResult readResult)
    {
        var stringValue = readResult.GetString();

        try
        {
            var uri = new Uri(stringValue!, UriKind.Absolute);
            return CellMapperResult.Success(uri);
        }
        catch (Exception exception)
        {
            return CellMapperResult.Invalid(exception);
        }
    }
}

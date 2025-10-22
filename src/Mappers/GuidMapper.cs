namespace ExcelMapper.Mappers;

/// <summary>
/// A mapper that tries to map the value of a cell to a <see cref="Guid"/>.
/// </summary>
public class GuidMapper : ICellMapper
{
    public CellMapperResult Map(ReadCellResult readResult)
    {
        var stringValue = readResult.GetString();

        try
        {
            var result = Guid.Parse(stringValue!);
            return CellMapperResult.Success(result);
        }
        catch (Exception exception)
        {
            return CellMapperResult.Invalid(exception);
        }
    }
}

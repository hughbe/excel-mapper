namespace ExcelMapper.Mappers;

/// <summary>
/// A mapper that tries to map the value of a cell to a <see cref="Guid"/>.
/// </summary>
public class GuidMapper : ICellMapper
{
    /// <inheritdoc/>
    public CellMapperResult Map(ReadCellResult readResult)
    {
        var stringValue = readResult.GetString();

        try
        {
            if (Guid.TryParse(stringValue, out var result))
            {
                return CellMapperResult.Success(result);
            }
        }
        catch (Exception exception)
        {
            return CellMapperResult.Invalid(exception);
        }

        // Only create the exception (by parsing again) if it is needed.
        return MapperUtilities.InvalidFromParse(() => Guid.Parse(stringValue!));
    }
}

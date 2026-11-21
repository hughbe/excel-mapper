namespace ExcelMapper.Mappers;

/// <summary>
/// A mapper that tries to map the value of a cell to a <see cref="Guid"/>.
/// </summary>
internal class GuidMapper : ICellMapper, ITypedCellMapper<Guid>, ITypedCellMapper<Guid?>
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
        return InvalidResult(stringValue);
    }

    CellConversion ITypedCellMapper<Guid>.Convert(ReadCellResult cell, out Guid value)
    {
        try
        {
            return Guid.TryParse(cell.GetString(), out value) ? CellConversion.Success : CellConversion.Invalid;
        }
        catch
        {
            value = default;
            return CellConversion.Invalid;
        }
    }

    CellConversion ITypedCellMapper<Guid?>.Convert(ReadCellResult cell, out Guid? value)
    {
        var conversion = ((ITypedCellMapper<Guid>)this).Convert(cell, out var result);
        value = conversion == CellConversion.Invalid ? null : result;
        return conversion;
    }

    // A separate method so that the closure capturing these values is only allocated for invalid
    // values. Creating the lambda in Map would allocate the closure for every cell.
    private static CellMapperResult InvalidResult(string? stringValue)
        => MapperUtilities.InvalidFormat(stringValue, "a valid Guid", () => Guid.Parse(stringValue!));
}

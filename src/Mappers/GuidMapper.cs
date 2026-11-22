namespace ExcelMapper.Mappers;

internal sealed class GuidMapper : ICellMapper<Guid>, ICellMapper<Guid?>
{
    CellConversion ICellMapper<Guid>.Convert(ReadCellResult cell, out Guid value)
        => Guid.TryParse(cell.GetString(), out value) ? CellConversion.Success : CellConversion.Invalid;

    CellConversion ICellMapper<Guid?>.Convert(ReadCellResult cell, out Guid? value)
    {
        var conversion = ((ICellMapper<Guid>)this).Convert(cell, out var result);
        value = conversion == CellConversion.Invalid ? null : result;
        return conversion;
    }

    public Exception CreateException(ReadCellResult cell)
    {
        var stringValue = cell.GetString();
        return MapperUtilities.InvalidFormat(stringValue, "a valid Guid", () => Guid.Parse(stringValue!));
    }
}

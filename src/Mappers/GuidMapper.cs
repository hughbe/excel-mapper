namespace ExcelMapper.Mappers;

internal sealed class GuidMapper : ICellMapper<Guid>, ICellMapper<Guid?>
{
    CellConversion ICellMapper<Guid>.Convert(ReadCellResult cell, in CellContext context, out Guid value, ref object? error)
        => Guid.TryParse(cell.GetString(), out value) ? CellConversion.Success : CellConversion.Invalid;

    CellConversion ICellMapper<Guid?>.Convert(ReadCellResult cell, in CellContext context, out Guid? value, ref object? error)
    {
        var conversion = ((ICellMapper<Guid>)this).Convert(cell, in context, out var result, ref error);
        value = conversion == CellConversion.Invalid ? null : result;
        return conversion;
    }

    public Exception CreateException(ReadCellResult cell, in CellContext context, object? error)
    {
        var stringValue = cell.GetString();
        return MapperUtilities.InvalidFormat(stringValue, "a valid Guid", () => Guid.Parse(stringValue!));
    }
}

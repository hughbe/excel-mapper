namespace ExcelMapper.Mappers;

internal sealed class VersionMapper : ICellMapper<Version?>
{
    public CellConversion Convert(ReadCellResult cell, in CellContext context, out Version? value, ref object? error)
        => Version.TryParse(cell.GetString(), out value) ? CellConversion.Success : CellConversion.Invalid;

    public Exception CreateException(ReadCellResult cell, object? error)
    {
        var stringValue = cell.GetString();
        return MapperUtilities.InvalidFromParse(() => new Version(stringValue!));
    }
}

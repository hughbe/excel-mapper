namespace ExcelMapper.Mappers;

internal sealed class VersionMapper : ICellMapper<Version?>
{
    public CellConversion Convert(ReadCellResult cell, out Version? value)
        => Version.TryParse(cell.GetString(), out value) ? CellConversion.Success : CellConversion.Invalid;

    public Exception CreateException(ReadCellResult cell)
    {
        var stringValue = cell.GetString();
        return MapperUtilities.InvalidFromParse(() => new Version(stringValue!));
    }
}

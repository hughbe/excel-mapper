namespace ExcelMapper.Mappers;

internal sealed class VersionMapper : ICellMapper<Version?>
{
    /// <summary>
    /// The mapper, shared by every member: it holds no state.
    /// </summary>
    public static VersionMapper Instance { get; } = new();

    private VersionMapper()
    {
    }

    public CellConversion Convert(ReadCellResult cell, in CellContext context, out Version? value, ref object? error)
        => Version.TryParse(cell.GetString(), out value) ? CellConversion.Success : CellConversion.Invalid;

    public Exception CreateException(ReadCellResult cell, in CellContext context, object? error)
    {
        var stringValue = cell.GetString();
        return MapperUtilities.InvalidFromParse(() => new Version(stringValue!));
    }
}

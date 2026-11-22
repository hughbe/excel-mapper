namespace ExcelMapper.Mappers;

internal sealed class UriMapper : ICellMapper<Uri?>
{
    /// <summary>
    /// The mapper, shared by every member: the kind of URI comes from the cell's
    /// <see cref="CellContext.Options"/>.
    /// </summary>
    public static UriMapper Instance { get; } = new();

    private UriMapper()
    {
    }

    public CellConversion Convert(ReadCellResult cell, in CellContext context, out Uri? value, ref object? error)
    {
        var stringValue = cell.GetString();
        if (stringValue is not null && Uri.TryCreate(stringValue, context.Options.UriKind ?? UriKind.Absolute, out value))
        {
            return CellConversion.Success;
        }

        value = null;
        return CellConversion.Invalid;
    }

    public Exception CreateException(ReadCellResult cell, in CellContext context, object? error)
    {
        var uriKind = context.Options.UriKind ?? UriKind.Absolute;
        return new UriFormatException($"The value \"{cell.GetString()}\" isn't a valid {(uriKind == UriKind.RelativeOrAbsolute ? "" : uriKind.ToString().ToLowerInvariant() + " ")}URI.");
    }
}

namespace ExcelMapper.Mappers;

internal sealed class UriMapper : ICellMapper<Uri?>
{
    private UriKind _uriKind = UriKind.Absolute;

    public UriKind UriKind
    {
        get => _uriKind;
        set
        {
            EnumUtilities.ValidateIsDefined(value);
            _uriKind = value;
        }
    }

    public CellConversion Convert(ReadCellResult cell, in CellContext context, out Uri? value, ref object? error)
    {
        var stringValue = cell.GetString();
        if (stringValue is not null && Uri.TryCreate(stringValue, UriKind, out value))
        {
            return CellConversion.Success;
        }

        value = null;
        return CellConversion.Invalid;
    }

    public Exception CreateException(ReadCellResult cell, object? error)
    {
        var uriKind = UriKind;
        return new UriFormatException($"The value \"{cell.GetString()}\" isn't a valid {(uriKind == UriKind.RelativeOrAbsolute ? "" : uriKind.ToString().ToLowerInvariant() + " ")}URI.");
    }
}

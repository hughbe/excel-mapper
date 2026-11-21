namespace ExcelMapper.Mappers;

/// <summary>
/// Tries to map the value of a cell to a <see cref="Uri"/>.
/// </summary>
public class UriMapper : ICellMapper
{
    private UriKind _uriKind = UriKind.Absolute;

    /// <summary>
    /// Gets or sets the kind of URI to create.
    /// </summary>
    public UriKind UriKind
    {
        get => _uriKind;
        set
        {
            EnumUtilities.ValidateIsDefined(value);
            _uriKind = value;
        }
    }

    /// <inheritdoc/>
    public CellMapperResult Map(ReadCellResult readResult)
    {
        var stringValue = readResult.GetString();

        var uriKind = UriKind;
        if (stringValue is not null && Uri.TryCreate(stringValue, uriKind, out var uri))
        {
            return CellMapperResult.Success(uri);
        }

        // Only create the exception (by parsing again) if it is needed.
        return InvalidResult(stringValue, uriKind);
    }

    // A separate method so that the closure capturing these values is only allocated for invalid
    // values. Creating the lambda in Map would allocate the closure for every cell.
    private static CellMapperResult InvalidResult(string? stringValue, UriKind uriKind)
        => stringValue is null
            ? MapperUtilities.InvalidFromParse(() => new Uri(stringValue!, uriKind))
            : CellMapperResult.Invalid(() => new UriFormatException($"The value \"{stringValue}\" isn't a valid {(uriKind == UriKind.RelativeOrAbsolute ? "" : uriKind.ToString().ToLowerInvariant() + " ")}URI."));
}

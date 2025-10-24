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

        try
        {
            var uri = new Uri(stringValue!, UriKind);
            return CellMapperResult.Success(uri);
        }
        catch (Exception exception)
        {
            return CellMapperResult.Invalid(exception);
        }
    }
}

namespace ExcelMapper.Mappers;

/// <summary>
/// Reads a cell's text. A later mapper that converts the value, such as a mapping dictionary,
/// replaces it.
/// </summary>
internal sealed class StringMapper : ICellMapper<string?>
{
    public CellConversion Convert(ReadCellResult cell, out string? value)
    {
        value = cell.GetString();
        return CellConversion.SuccessUnlessReplaced;
    }

    // Text is never invalid.
    public Exception CreateException(ReadCellResult cell) => new FormatException("The value couldn't be read as text.");
}

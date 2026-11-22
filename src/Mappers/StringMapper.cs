namespace ExcelMapper.Mappers;

/// <summary>
/// Reads a cell's text. A later mapper that converts the value, such as a mapping dictionary,
/// replaces it.
/// </summary>
internal sealed class StringMapper : ICellMapper<string?>
{
    /// <summary>
    /// The mapper, shared by every member: it holds no state.
    /// </summary>
    public static StringMapper Instance { get; } = new();

    private StringMapper()
    {
    }

    public CellConversion Convert(ReadCellResult cell, in CellContext context, out string? value, ref object? error)
    {
        value = cell.GetString();
        return CellConversion.SuccessUnlessReplaced;
    }

    // Text is never invalid.
    public Exception CreateException(ReadCellResult cell, in CellContext context, object? error) => new FormatException("The value couldn't be read as text.");
}

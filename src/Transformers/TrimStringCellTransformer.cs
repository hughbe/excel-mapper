namespace ExcelMapper.Transformers;

/// <summary>
/// Trims text, for <see cref="ExcelTrimStringAttribute"/> and <c>WithTrim()</c>.
/// </summary>
internal sealed class TrimStringCellTransformer : ICellTransformer
{
    public static TrimStringCellTransformer Instance { get; } = new();

    private TrimStringCellTransformer()
    {
    }

    // Trimming only changes text. Other values, such as dates and numbers, keep their type rather
    // than being converted to text and parsed again. With formatting preserved, every value is text.
    public CellValue Transform(CellValue value, CellContext context)
        => value.Value is string text ? new CellValue(text.Trim()) : value;
}
